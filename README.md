# PE Warranty Form — Frontend (User Form Page)

A React-based single-page application (SPA) for **True Sun Trading Company** that allows solar installation integrators to submit Premier Energies (PE) warranty certificate requests and track the review status of their submission.

---

## Table of Contents

- [Project Overview](#project-overview)
- [Tech Stack](#tech-stack)
- [Architecture](#architecture)
- [Project Structure](#project-structure)
- [Environment Variables](#environment-variables)
- [Getting Started](#getting-started)
- [Routing / Navigation](#routing--navigation)
- [Component Reference](#component-reference)
  - [App.js](#appjs)
  - [Home.jsx](#homejsx)
  - [Form.jsx](#formjsx)
  - [TrackRequest.jsx](#trackrequestjsx)
- [Utilities](#utilities)
  - [cloudinary.js](#cloudinaryjs)
  - [firebase.js](#firebasejs)
- [API Integration](#api-integration)
- [Form Validation Rules](#form-validation-rules)
- [Data Flow](#data-flow)
- [Build & Deployment](#build--deployment)

---

## Project Overview

This app serves as the **customer-facing frontend** for a warranty verification workflow:

1. **Integrators** (solar installation companies/EPCs) visit the site and fill out a warranty request form.
2. The request is saved to a backend (Express.js on Render) and images are uploaded to Cloudinary.
3. **Admin staff** at True Sun Trading review the request on a separate admin panel and update its status.
4. **Integrators** can return to the site and enter their Request ID to check the decision status.

---

## Tech Stack

| Layer | Technology |
|---|---|
| UI Framework | React 19 (Create React App) |
| Styling | Tailwind CSS v3 |
| Icons | Lucide React |
| Excel Parsing | SheetJS (`xlsx`) |
| Image Hosting | Cloudinary (via unsigned upload preset) |
| Backend API | Express.js on Render (`pe-warranty-backend.onrender.com`) |
| Database | Firebase Firestore (managed by backend) |

> **Note:** `firebase.js` is present in the client but the Firestore SDK is **not used directly** — all database reads/writes are proxied through the Express backend API. The file is kept for potential future use.

---

## Architecture

```
Browser (React SPA)
       │
       ├── URL query-string routing (?form, ?track)
       │
       ├── Form.jsx ──► POST /api/requests    ──► Express Backend ──► Firestore
       │             ──► PUT  /api/requests/:id
       │             ──► Cloudinary (image upload, direct from browser)
       │
       └── TrackRequest.jsx ──► GET /api/requests/:id ──► Express Backend ──► Firestore
```

---

## Project Structure

```
PE-Warranty-form/
├── public/                   # Static assets served as-is
├── src/
│   ├── assets/
│   │   └── Images/
│   │       ├── Trusunlogo.png   # True Sun Trading Company logo
│   │       └── premier.png      # Premier Energies logo
│   ├── components/
│   │   ├── Home.jsx             # Landing page with navigation cards
│   │   ├── Form.jsx             # Warranty request form (main component)
│   │   └── TrackRequest.jsx     # Request status tracker
│   ├── utils/
│   │   └── cloudinary.js        # Cloudinary unsigned image upload helper
│   ├── App.js                   # Root component — URL-based view router
│   ├── firebase.js              # Firebase/Firestore initialisation (unused on client)
│   ├── index.js                 # React DOM entry point
│   ├── index.css                # Global base styles
│   └── App.css                  # App-level animation utilities
├── tailwind.config.js           # Tailwind CSS configuration
├── package.json
└── .env                         # Environment variables (NOT committed to git)
```

---

## Environment Variables

Create a `.env` file in the project root with the following variables. **Never commit real values to version control.**

```env
# Cloudinary — for direct browser-to-cloud image uploads
REACT_APP_CLOUDINARY_CLOUD_NAME=your_cloud_name
REACT_APP_CLOUDINARY_UPLOAD_PRESET=your_unsigned_upload_preset

# Firebase — initialisation config (even if not used directly on the client)
REACT_APP_FIREBASE_API_KEY=your_api_key
REACT_APP_FIREBASE_AUTH_DOMAIN=your_project_id.firebaseapp.com
REACT_APP_FIREBASE_PROJECT_ID=your_project_id
REACT_APP_FIREBASE_STORAGE_BUCKET=your_project_id.appspot.com
REACT_APP_FIREBASE_MESSAGING_SENDER_ID=your_sender_id
REACT_APP_FIREBASE_APP_ID=your_app_id
```

> All environment variable names **must** start with `REACT_APP_` for Create React App to expose them to the browser bundle.

---

## Getting Started

### Prerequisites

- Node.js ≥ 18
- npm ≥ 9

### Installation

```bash
# Clone the repository
git clone <repo-url>
cd PE-Warranty-form

# Install dependencies
npm install

# Create your environment file
cp .env.example .env   # then edit .env with your real values

# Start the development server
npm start
# → Opens http://localhost:3000
```

### Production Build

```bash
npm run build
# Output is placed in /build — serve this with any static host (Netlify, Vercel, S3, etc.)
```

---

## Routing / Navigation

This app uses **URL query-string routing** instead of React Router. The `App.js` component reads `window.location.search` on mount and renders the correct view:

| URL | View Rendered |
|---|---|
| `/` | `Home.jsx` (landing page) |
| `/?form=true` | `Form.jsx` (new warranty request) |
| `/?form&edit=WR_1234` | `Form.jsx` in **edit mode** (pre-filled with existing request data) |
| `/?track=true` | `TrackRequest.jsx` (status tracker) |

Navigation between views is handled by directly setting `window.location.href`.

---

## Component Reference

### `App.js`

**Root component.** Handles two responsibilities on mount:

1. **Server wake-up ping** — Because the backend runs on Render's free tier, it spins down after inactivity. A `fetch` to the backend root endpoint is made immediately on app load to wake it up before the user interacts.
2. **URL-based routing** — Reads query params to set the `view` state (`'home'` | `'form'` | `'track'`).

```
State:
  view  string  — controls which page component is rendered
```

---

### `Home.jsx`

**Landing page.** Displays the True Sun Trading Company branding and two navigation cards:

| Card | Action |
|---|---|
| **New Request** | Navigates to `/?form=true` |
| **Track Status** | Navigates to `/?track=true` |

Decorative animated blob shapes are rendered as absolutely-positioned `div`s using Tailwind `animate-blob` utilities defined in `App.css`.

---

### `Form.jsx`

The **main and largest component** (~915 lines). Handles the entire warranty submission workflow.

#### Sub-components (defined in the same file)

| Component | Purpose |
|---|---|
| `Input` | Reusable controlled text input with label and optional tooltip. Defined at module scope (outside `FormPage`) to prevent re-mounting on each render. |
| `ConfirmationModal` | Full-screen overlay modal asking the user to confirm submission. |

#### State

| State Variable | Type | Description |
|---|---|---|
| `formData` | object | All text fields of the form (integrator + customer details) |
| `serialNumbers` | string[] | Dynamic list of solar panel serial numbers |
| `files` | File[] | New image files selected by the user (to be uploaded) |
| `previewImages` | string[] | Blob URLs (new) or Cloudinary HTTPS URLs (existing) for image previews |
| `requestId` | string | The document ID returned by the backend after successful submission |
| `isSubmitted` | boolean | Whether the form has been successfully submitted (shows success screen) |
| `isSubmitting` | boolean | Locks UI during async submission |
| `showConfirmation` | boolean | Controls the `ConfirmationModal` visibility |
| `editModeId` | string\|null | If present, the form operates in edit/resubmit mode for this request ID |
| `copied` | boolean | Tracks clipboard copy state for the Request ID button |

#### Key Handlers

| Handler | Description |
|---|---|
| `handleChange(e)` | Generic field change handler. Applies field-specific sanitisation/validation (numbers only for phone/pincode, restricted chars for email/address/district). |
| `addSerialNumber()` | Appends a new empty entry to the `serialNumbers` array. |
| `updateSerialNumber(index, value)` | Updates a single serial number. Validates: no whitespace/commas, ≤25 chars, no duplicates. |
| `removeSerialNumber(index)` | Removes a serial number entry (minimum 1 must remain). |
| `handleFileChange(e)` | Adds selected image files for upload. Enforces a max of **2 site pictures**. |
| `removeImage(index)` | Removes an image from both `previewImages` and `files`. |
| `handleExcelUpload(e)` | Parses an `.xlsx` / `.xls` file using SheetJS, extracts serial numbers from column A, deduplicates, and merges into the existing list. |
| `handleSubmit(e)` | Pre-submission validation (exactly 2 images, valid Cloudinary preset). Opens the `ConfirmationModal`. |
| `processSubmission()` | The actual async submission logic: uploads images to Cloudinary, then POSTs or PUTs to the backend. |
| `handleReset()` | Resets all form state to allow a new submission. |
| `fetchRequestData(id)` | On edit mode, fetches existing request data from the backend and pre-fills the form. Parses the composite address strings back into individual location/district/pincode fields. |
| `handleCopy()` | Copies the Request ID to clipboard and shows a 2-second "Copied!" feedback. |

#### Address Field Format

The form breaks addresses into **3 separate fields** (Location, District, Pincode) for usability. Before sending to the backend, they are merged into a single string:

```
Office: "{officeLocation}, {officeDistrict} - {officePincode}"
Site:   "{customerName}, {siteLocation}, {siteDistrict} - {sitePincode}"
```

When fetching an existing request for editing, `fetchRequestData` reverses this by parsing the combined strings using `splitAddress()`.

---

### `TrackRequest.jsx`

**Status tracker page.** Allows a user to look up their request by ID.

#### State

| State Variable | Type | Description |
|---|---|---|
| `requestId` | string | ID entered by the user in the search input |
| `status` | string\|null | Status value returned from the backend API |
| `rejectionReason` | string | Optional reason shown when a request is rejected |
| `loading` | boolean | Controls the search button spinner |
| `error` | string | Error message to display if the ID is not found |
| `searched` | boolean | Whether a search has been attempted (controls result visibility) |

#### Status Display Logic (`getStatusConfig`)

Maps the raw status string from the backend to a UI config object:

| Backend Status Values | Displayed As |
|---|---|
| `accepted`, `approved`, `verified`, `completed` | ✅ Green — Accepted |
| `rejected`, `declined`, `denied` | ❌ Red — Rejected (+ Edit & Resubmit button) |
| `pending`, `in review`, `processing` | 🕐 Amber — Pending Review |
| Any other value | 🔵 Blue — Unknown Status |

When a request is **rejected**, an **"Edit & Resubmit"** button appears, which navigates to `/?form&edit={requestId}`.

---

## Utilities

### `cloudinary.js`

**`uploadToCloudinary(file: File): Promise<string>`**

Uploads a single image file directly from the browser to Cloudinary using an **unsigned upload preset** (no backend involvement). Returns the `secure_url` of the uploaded image.

- Reads `REACT_APP_CLOUDINARY_CLOUD_NAME` and `REACT_APP_CLOUDINARY_UPLOAD_PRESET` from env.
- Throws if configuration is missing or if the upload fails.

### `firebase.js`

Initialises a Firebase app using env-based config and exports a `db` (Firestore) instance. **This `db` export is not currently used in any component** — all data operations go through the Express backend. It is preserved for potential future direct client-side Firestore access.

---

## API Integration

All API calls target the base URL: `https://pe-warranty-backend.onrender.com`

| Method | Endpoint | Used In | Description |
|---|---|---|---|
| `GET` | `/` | `App.js` | Wake-up ping on app load |
| `GET` | `/api/requests/:id` | `TrackRequest.jsx`, `Form.jsx` | Fetch request by ID (status + full data) |
| `POST` | `/api/requests` | `Form.jsx` | Create a new warranty request |
| `PUT` | `/api/requests/:id` | `Form.jsx` | Update an existing (rejected) request |

**Request Body for POST/PUT:**

```json
{
  "integratorName": "string",
  "officeAddress": "Location, District - Pincode",
  "contactPerson": "string",
  "contactNo": "string",
  "email": "string",
  "customerProjectSite": "CustomerName, Location, District - Pincode",
  "customerContact": "string",
  "customerAlternate": "string",
  "customerEmail": "string",
  "customerAlternateEmail": "string",
  "serialNumbers": ["SN-001", "SN-002"],
  "sitePictures": ["https://res.cloudinary.com/..."]
}
```

**Response Body (success):**

```json
{ "id": "WR_1234" }
```

---

## Form Validation Rules

| Field | Rule |
|---|---|
| `contactNo` | Digits only |
| `officePincode`, `sitePincode` | Digits only, max 6 digits |
| `email`, `customerEmail` | Alphanumeric + `@`, `.`, `_`, `-` only |
| `officeLocation`, `siteLocation` | Letters, numbers, spaces, `.`, `,`, `-`, `/`, `#` only |
| `officeDistrict`, `siteDistrict` | Letters and spaces only |
| Serial Numbers | No whitespace or commas; max 25 characters; no duplicates |
| Site Pictures | Exactly **2** images required to submit |
| Excel Upload | Reads column A; skips invalid/duplicate entries; alerts on filtered count |

---

## Data Flow

```
User fills form
      │
      ▼
handleSubmit() ──► Validates (2 images, env config)
      │
      ▼
ConfirmationModal shown
      │ user clicks "Yes, Submit"
      ▼
processSubmission()
  ├─ 1. Upload images ──► Cloudinary API ──► returns secure_url[]
  └─ 2. POST/PUT ──────► Express Backend ──► Firestore ──► returns { id }
      │
      ▼
isSubmitted = true → Shows success screen with Request ID
```

---

## Build & Deployment

```bash
npm run build
```

The `/build` directory is a production-ready static bundle. Deploy to any static hosting provider (Netlify, Vercel, GitHub Pages, Firebase Hosting, AWS S3 + CloudFront, etc.).

**Important:** Set all `REACT_APP_*` environment variables in your hosting provider's dashboard, as the `.env` file is local only.
