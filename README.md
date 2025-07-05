# Auto Medfill – Microsoft Word Add-in

A custom MS Word Add-in that intelligently suggests medicine names and dosages via a dynamic dropdown in a custom task pane. Designed to speed up medical presription writing using real-time text selection and contextual autofill.

---

## Features

- Custom **task pane** with dynamic UI powered by Fluent UI.
- **Dropdown suggestions** for medicine names based on selected or typed text.
- Built using **Office.js** API to interact with Word document objects.
- Works on Word for the Web.

---

## Tech Stack

- JavaScript
- HTML, CSS
- [Office.js](https://learn.microsoft.com/javascript/api/overview/office) (Word API)
- Fluent UI (Fabric CSS)
- Node.js + npm (for development/build tooling)
- Webpack (for bundling)

---

## Getting Started

### Prerequisites

- Node.js (v14 or higher)

### Installation

```bash
npm install
npm start
```
<br>
This will:

- Start a **local development server** (At `https://localhost:3000`).
- Bundle the frontend code and host it for Word to use.

---

## Sideloading the Add-in

To load the add-in into Microsoft Word:

1. Run `npm start` to launch the dev server.  
2. Open Microsoft Word.  
3. Go to **Insert > Add-ins > My Add-ins > Upload My Add-in**.  
4. Choose the `manifest.xml` file from the project root.  
5. The task pane will load your local add-in.
<br>
The `manifest.xml` file tells Word how to find and load the add-in.
