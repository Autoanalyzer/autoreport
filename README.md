Auto Report — pH Calibration Certificate

Overview
- Static web app (HTML/CSS/JS) to fill a form and export a pH calibration certificate as a PDF.
- Lives entirely in the browser. No server or install needed.

How to use
1) Open `index.html` with any modern browser (Edge/Chrome).
2) Fill fields on the left. The right side updates live.
3) Add more buffer rows or calibration points as needed.
4) Upload a logo and up to 3 photos (optional).
5) Click `บันทึกเป็น PDF` to export. The file name uses Tag No + date.

Tips
- Use `บันทึก JSON` to save your inputs as a `.json` file and `นำเข้า JSON` to load them later.
- Data also auto-saves to `localStorage` and reloads on next visit.
- The app loads `html2pdf.bundle.min.js` from a CDN. If you need offline-only usage without Internet, download the bundle and place it at `lib/html2pdf.bundle.min.js` and update the `<script>` tag in `index.html` accordingly.

Customize
- Update layout or styles in `assets/style.css`.
- Add/rename fields in `index.html` and wire them in `assets/script.js` (`state` object and bindings).

Notes
- PDF rendering uses HTML/CSS; final output may differ slightly from Word templates. Adjust CSS as needed for your organization’s look.

