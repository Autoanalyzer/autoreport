# pH Module Customization

This folder isolates pH-specific customizations so you can edit safely
without touching core files like `assets/style.css` or `assets/script.js`.

Edit these files:

- `assets/ph/custom.css` – CSS overrides for layout/appearance
- `assets/ph/custom.js` – Small DOM tweaks, defaults, and behaviors

Tips
- Keep selectors specific (prefix with `.cert`, `.cert-header`, etc.) so
  overrides apply only to the certificate UI.
- Prefer overrides here rather than editing the base files directly.
- If you need a new hook, add it here first; if it grows large we can split
  into smaller modules under `assets/ph/`.
