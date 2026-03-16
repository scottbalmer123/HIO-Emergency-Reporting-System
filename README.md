# Mine Emergency Incident Reports

Mine emergency incident reporting app for Hancock Iron Ore. The app now supports two runtime modes:

- `Google Cloud mode`: the UI is served by Cloud Run, structured data is stored in Firestore, and incident photos are stored in Cloud Storage.
- `Local fallback mode`: if the cloud API is unavailable, the browser falls back to `localStorage`.

## Architecture

- Front end: static SPA in [index.html](/Users/scottbalmer/Documents/Playground/index.html), [app.js](/Users/scottbalmer/Documents/Playground/app.js), and [style.css](/Users/scottbalmer/Documents/Playground/style.css)
- Backend: Node/Express service in [server.js](/Users/scottbalmer/Documents/Playground/server.js) using the Firebase Admin SDK for Firestore and Storage access
- Structured storage: Firestore
- Photo storage: Cloud Storage
- Deployment target: Cloud Run

This split is important because incident photos should not be stored inline in Firestore documents. The app now stores photo metadata in Firestore and the image objects themselves in Cloud Storage.

## Features

- Multi-view workflow with `New Incident`, `Incident Register`, `Reporting`, `Personnel`, and `Settings`
- Incident intake form with personnel assignment, radio timeline, and photo uploads
- Reporting dashboard with filterable analytics and CSV export
- HIO-style PDF export
- Cloud-aware persistence with same-origin `/api/*` endpoints
- Storage status indicator in the header so operators can see whether the app is connected to Google Cloud or running in local fallback mode

## Local Development

Install dependencies:

```bash
npm install
```

Run the Cloud Run-compatible server locally:

```bash
npm start
```

Then open [http://localhost:8080](http://localhost:8080).

If Google Cloud credentials and environment variables are not configured, the UI still runs and falls back to browser storage.

## Google Cloud Deployment

### 1. Prepare Google Cloud resources

- Create a Firestore database in `Native` mode.
- Create a Cloud Storage bucket for incident photos.
- Create or choose a Cloud Run service account.
- Grant that service account access to Firestore and object read/write access to the bucket.

### 2. Configure environment variables

Use the variables shown in [.env.example](/Users/scottbalmer/Documents/Playground/.env.example):

- `APP_NAMESPACE`: logical environment name such as `production`
- `FIRESTORE_COLLECTION`: top-level Firestore collection root for this app
- `FIREBASE_STORAGE_BUCKET`: bucket used for incident photos
- `FIREBASE_SERVICE_ACCOUNT_PATH`: path to a Firebase service account JSON file
- `FIREBASE_SERVICE_ACCOUNT_JSON`: optional inline JSON alternative for secret injection

`PORT` is injected automatically by Cloud Run.

If you do not provide an explicit Firebase service account, the server falls back to application default credentials, which is the preferred Cloud Run path.

### 3. Enable required APIs

```bash
gcloud services enable \
  run.googleapis.com \
  cloudbuild.googleapis.com \
  artifactregistry.googleapis.com \
  firestore.googleapis.com \
  storage.googleapis.com
```

### 4. Deploy to Cloud Run

Example:

```bash
gcloud run deploy hio-incident-control \
  --source . \
  --region australia-southeast1 \
  --allow-unauthenticated \
  --set-env-vars APP_NAMESPACE=production,FIRESTORE_COLLECTION=mineEmergencyApp,FIREBASE_STORAGE_BUCKET=YOUR_BUCKET_NAME
```

If you are using a dedicated service account, add:

```bash
--service-account YOUR_SERVICE_ACCOUNT
```

## API Summary

- `GET /api/bootstrap`: load settings, personnel, and incidents
- `PUT /api/settings`: save radio/unit settings
- `PUT /api/personnel`: save the personnel roster
- `PUT /api/incidents/:incidentId`: save one incident
- `DELETE /api/incidents/:incidentId`: remove a cloud incident record
- `POST /api/incidents/bulk`: bulk sync incidents
- `GET /api/assets/:incidentId/:photoId`: stream an incident photo from Cloud Storage
- `GET /api/health`: health check

## Google Cloud Optimisations Included

- Cloud Run serves both the UI and API on the same origin, so there is no CORS layer to manage.
- Incidents are stored as individual Firestore documents instead of one giant blob, which is safer for growth and updates.
- Incident photos are streamed from Cloud Storage instead of being embedded directly in Firestore.
- The browser keeps a local cache for resilience, but the intended source of truth in production is Google Cloud.

## Files Added For Cloud Deployment

- [server.js](/Users/scottbalmer/Documents/Playground/server.js)
- [package.json](/Users/scottbalmer/Documents/Playground/package.json)
- [Dockerfile](/Users/scottbalmer/Documents/Playground/Dockerfile)
- [.dockerignore](/Users/scottbalmer/Documents/Playground/.dockerignore)
- [.gcloudignore](/Users/scottbalmer/Documents/Playground/.gcloudignore)
- [.env.example](/Users/scottbalmer/Documents/Playground/.env.example)
