const express = require("express");
const path = require("path");
const admin = require("firebase-admin");

const PORT = Number(process.env.PORT || 8080);
const APP_NAMESPACE = process.env.APP_NAMESPACE || "default";
const FIRESTORE_COLLECTION = process.env.FIRESTORE_COLLECTION || "mineEmergencyApp";
const FIREBASE_STORAGE_BUCKET = process.env.FIREBASE_STORAGE_BUCKET || process.env.GCS_BUCKET_NAME || "";
const FIREBASE_SERVICE_ACCOUNT_PATH = process.env.FIREBASE_SERVICE_ACCOUNT_PATH || process.env.GOOGLE_APPLICATION_CREDENTIALS || "";
const FIREBASE_SERVICE_ACCOUNT_JSON = process.env.FIREBASE_SERVICE_ACCOUNT_JSON || "";
const STATIC_FILES = new Map([
  ["/", "index.html"],
  ["/index.html", "index.html"],
  ["/app.js", "app.js"],
  ["/style.css", "style.css"],
  ["/logo-hancock-iron-ore.svg", "logo-hancock-iron-ore.svg"],
]);

const app = express();

function resolveFirebaseCredential() {
  if (FIREBASE_SERVICE_ACCOUNT_JSON) {
    return admin.credential.cert(JSON.parse(FIREBASE_SERVICE_ACCOUNT_JSON));
  }

  if (FIREBASE_SERVICE_ACCOUNT_PATH) {
    return admin.credential.cert(require(path.resolve(FIREBASE_SERVICE_ACCOUNT_PATH)));
  }

  return admin.credential.applicationDefault();
}

if (!admin.apps.length) {
  const firebaseOptions = {
    credential: resolveFirebaseCredential(),
  };

  if (FIREBASE_STORAGE_BUCKET) {
    firebaseOptions.storageBucket = FIREBASE_STORAGE_BUCKET;
  }

  admin.initializeApp(firebaseOptions);
}

const firestore = admin.firestore();
const bucket = FIREBASE_STORAGE_BUCKET ? admin.storage().bucket(FIREBASE_STORAGE_BUCKET) : null;

app.use(express.json({ limit: "25mb" }));

function asyncHandler(handler) {
  return (req, res, next) => {
    Promise.resolve(handler(req, res, next)).catch(next);
  };
}

function appRoot() {
  return firestore.collection(FIRESTORE_COLLECTION).doc(APP_NAMESPACE);
}

function incidentsCollection() {
  return appRoot().collection("incidents");
}

function metaDocument(id) {
  return appRoot().collection("meta").doc(id);
}

function sanitizeSegment(value) {
  return String(value || "")
    .trim()
    .replace(/[^a-zA-Z0-9._-]+/g, "-")
    .replace(/^-+|-+$/g, "") || "item";
}

function photoExtension(contentType = "") {
  switch (contentType) {
    case "image/png":
      return "png";
    case "image/webp":
      return "webp";
    case "image/gif":
      return "gif";
    case "image/jpeg":
    default:
      return "jpg";
  }
}

function parseDataUrl(dataUrl) {
  const match = String(dataUrl || "").match(/^data:([^;]+);base64,(.+)$/);

  if (!match) {
    throw new Error("Incident photo data was not a valid base64 data URL.");
  }

  return {
    contentType: match[1],
    buffer: Buffer.from(match[2], "base64"),
  };
}

function buildPhotoUrl(incidentId, photoId) {
  return `/api/assets/${encodeURIComponent(incidentId)}/${encodeURIComponent(photoId)}`;
}

function formatPhotoForClient(incidentId, photo = {}) {
  const storagePath = typeof photo.storagePath === "string" ? photo.storagePath : "";

  return {
    id: photo.id,
    name: photo.name || "Incident photo",
    type: photo.type || "image/jpeg",
    storagePath,
    url: storagePath ? buildPhotoUrl(incidentId, photo.id) : "",
  };
}

function formatIncidentForClient(incident = {}) {
  const incidentId = incident.incidentId || "";
  const incidentPhotos = Array.isArray(incident.incidentPhotos)
    ? incident.incidentPhotos.map((photo) => formatPhotoForClient(incidentId, photo))
    : [];

  return {
    ...incident,
    incidentPhotos,
  };
}

async function savePhotoToBucket(incidentId, photo) {
  if (!bucket) {
    throw new Error("FIREBASE_STORAGE_BUCKET or GCS_BUCKET_NAME must be configured to store incident photos.");
  }

  const { contentType, buffer } = parseDataUrl(photo.dataUrl);
  const resolvedType = photo.type || contentType || "image/jpeg";
  const storagePath = [
    "apps",
    sanitizeSegment(APP_NAMESPACE),
    "incidents",
    sanitizeSegment(incidentId),
    "photos",
    `${sanitizeSegment(photo.id)}.${photoExtension(resolvedType)}`,
  ].join("/");

  await bucket.file(storagePath).save(buffer, {
    resumable: false,
    validation: false,
    contentType: resolvedType,
    metadata: {
      cacheControl: "private, max-age=3600",
    },
  });

  return {
    id: photo.id,
    name: photo.name || "Incident photo",
    type: resolvedType,
    storagePath,
  };
}

async function syncIncidentPhotos(incident, previousIncident = null) {
  const previousPhotos = new Map(
    Array.isArray(previousIncident?.incidentPhotos)
      ? previousIncident.incidentPhotos.map((photo) => [photo.id, photo])
      : []
  );
  const nextPhotos = Array.isArray(incident.incidentPhotos) ? incident.incidentPhotos : [];
  const savedPhotos = [];

  for (const photo of nextPhotos) {
    if (!photo || !photo.id) {
      continue;
    }

    if (photo.dataUrl) {
      savedPhotos.push(await savePhotoToBucket(incident.incidentId, photo));
      continue;
    }

    const previousPhoto = previousPhotos.get(photo.id);
    const storagePath = photo.storagePath || previousPhoto?.storagePath || "";

    savedPhotos.push({
      id: photo.id,
      name: photo.name || previousPhoto?.name || "Incident photo",
      type: photo.type || previousPhoto?.type || "image/jpeg",
      storagePath,
    });
  }

  if (bucket) {
    const nextPhotoIds = new Set(savedPhotos.map((photo) => photo.id));

    await Promise.all(
      [...previousPhotos.values()].map(async (photo) => {
        if (!photo?.id || !photo.storagePath || nextPhotoIds.has(photo.id)) {
          return;
        }

        try {
          await bucket.file(photo.storagePath).delete({ ignoreNotFound: true });
        } catch {
          // Keep the incident save successful even if a stale photo could not be deleted.
        }
      })
    );
  }

  return savedPhotos;
}

async function saveIncidentRecord(incident) {
  if (!incident || typeof incident !== "object" || !incident.incidentId) {
    throw new Error("Incident payload was invalid.");
  }

  const incidentRef = incidentsCollection().doc(String(incident.incidentId));
  const previousSnapshot = await incidentRef.get();
  const previousIncident = previousSnapshot.exists ? previousSnapshot.data() : null;
  const incidentPhotos = await syncIncidentPhotos(incident, previousIncident);
  const record = {
    ...incident,
    incidentPhotos,
    updatedAt: incident.updatedAt || new Date().toISOString(),
  };

  await incidentRef.set(record, { merge: false });
  return formatIncidentForClient(record);
}

async function deleteIncidentRecord(incidentId) {
  const incidentRef = incidentsCollection().doc(String(incidentId));
  const snapshot = await incidentRef.get();

  if (!snapshot.exists) {
    return false;
  }

  const incident = snapshot.data();

  if (bucket) {
    await Promise.all(
      (incident.incidentPhotos || []).map(async (photo) => {
        if (!photo?.storagePath) {
          return;
        }

        try {
          await bucket.file(photo.storagePath).delete({ ignoreNotFound: true });
        } catch {
          // Ignore stale object cleanup errors so the document can still be removed.
        }
      })
    );
  }

  await incidentRef.delete();
  return true;
}

async function loadBootstrapPayload() {
  const [settingsSnapshot, personnelSnapshot, incidentsSnapshot] = await Promise.all([
    metaDocument("settings").get(),
    metaDocument("personnel").get(),
    incidentsCollection().get(),
  ]);

  return {
    storageMode: "cloud",
    settings: settingsSnapshot.exists ? settingsSnapshot.data().settings || {} : {},
    personnel: personnelSnapshot.exists ? personnelSnapshot.data().personnel || [] : [],
    incidents: incidentsSnapshot.docs.map((snapshot) => formatIncidentForClient(snapshot.data())),
  };
}

app.get("/api/health", asyncHandler(async (_req, res) => {
  await firestore.doc("_healthcheck/status").get().catch(() => null);
  res.json({
    ok: true,
    storage: "google-cloud",
    firestoreCollection: FIRESTORE_COLLECTION,
    namespace: APP_NAMESPACE,
    bucketConfigured: Boolean(bucket),
  });
}));

app.get("/api/bootstrap", asyncHandler(async (_req, res) => {
  res.json(await loadBootstrapPayload());
}));

app.put("/api/settings", asyncHandler(async (req, res) => {
  const settings = req.body?.settings;

  if (!settings || typeof settings !== "object") {
    res.status(400).json({ error: "Settings payload was invalid." });
    return;
  }

  await metaDocument("settings").set(
    {
      settings,
      updatedAt: new Date().toISOString(),
    },
    { merge: true }
  );

  res.json({ settings });
}));

app.put("/api/personnel", asyncHandler(async (req, res) => {
  const personnel = Array.isArray(req.body?.personnel) ? req.body.personnel : null;

  if (!personnel) {
    res.status(400).json({ error: "Personnel payload was invalid." });
    return;
  }

  const sortedPersonnel = [...personnel].sort((a, b) => String(a.fullName || "").localeCompare(String(b.fullName || "")));

  await metaDocument("personnel").set(
    {
      personnel: sortedPersonnel,
      updatedAt: new Date().toISOString(),
    },
    { merge: true }
  );

  res.json({ personnel: sortedPersonnel });
}));

app.put("/api/incidents/:incidentId", asyncHandler(async (req, res) => {
  const incident = req.body?.incident;

  if (!incident || String(incident.incidentId) !== String(req.params.incidentId)) {
    res.status(400).json({ error: "Incident payload was invalid." });
    return;
  }

  res.json({ incident: await saveIncidentRecord(incident) });
}));

app.delete("/api/incidents/:incidentId", asyncHandler(async (req, res) => {
  const deleted = await deleteIncidentRecord(req.params.incidentId);
  res.json({ deleted });
}));

app.post("/api/incidents/bulk", asyncHandler(async (req, res) => {
  const incidents = Array.isArray(req.body?.incidents) ? req.body.incidents : null;

  if (!incidents) {
    res.status(400).json({ error: "Bulk incident payload was invalid." });
    return;
  }

  const savedIncidents = [];
  for (const incident of incidents) {
    savedIncidents.push(await saveIncidentRecord(incident));
  }

  res.json({ incidents: savedIncidents });
}));

app.get("/api/assets/:incidentId/:photoId", asyncHandler(async (req, res) => {
  if (!bucket) {
    res.status(404).json({ error: "Cloud photo storage is not configured." });
    return;
  }

  const incidentSnapshot = await incidentsCollection().doc(String(req.params.incidentId)).get();
  if (!incidentSnapshot.exists) {
    res.status(404).json({ error: "Incident not found." });
    return;
  }

  const photo = (incidentSnapshot.data().incidentPhotos || []).find((item) => item.id === req.params.photoId);
  if (!photo?.storagePath) {
    res.status(404).json({ error: "Incident photo not found." });
    return;
  }

  const file = bucket.file(photo.storagePath);
  const [exists] = await file.exists();

  if (!exists) {
    res.status(404).json({ error: "Stored photo object was not found." });
    return;
  }

  res.setHeader("Content-Type", photo.type || "application/octet-stream");
  res.setHeader("Cache-Control", "private, max-age=3600");
  file.createReadStream().pipe(res);
}));

for (const [route, file] of STATIC_FILES.entries()) {
  app.get(route, (_req, res) => {
    res.sendFile(path.join(__dirname, file));
  });
}

app.get("/favicon.ico", (_req, res) => {
  res.status(204).end();
});

app.use((err, _req, res, _next) => {
  console.error(err);
  res.status(500).json({
    error: err.message || "Internal server error",
  });
});

app.listen(PORT, () => {
  console.log(`Mine emergency app listening on port ${PORT}`);
});
