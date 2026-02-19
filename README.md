# PDF Nova

Application web locale type iLovePDF, optimisee pour un usage rapide:

- Merge de plusieurs PDFs
- Split PDF en plusieurs fichiers (N pages par fichier)
- Extraction de pages ciblees
- Rotation globale ou partielle
- Watermark texte
- Conversion images -> PDF
- Compression PDF
- Suppression pages blanches (heuristique)
- OCR texte (PDF/Image -> TXT)
- Extraction video sociale en MP4 (YouTube/Twitter/TikTok via lien)
- Convertisseur: PDF -> DOCX, PDF -> images (ZIP), PDF -> Excel (XLSX), Image -> PDF, Office -> PDF

## Lancer

```powershell
cd pdf_nova
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python server.py
```

Puis ouvre:

- `http://127.0.0.1:8091`

## Lancer en app desktop (ouvre l'interface web PDF Nova automatiquement)

```powershell
cd pdf_nova
pip install -r requirements.txt
python desktop_launcher.py
```

## Notes

- Tout tourne en local.
- Le launcher desktop choisit un port libre automatiquement si `8091` est deja pris.
- Les fichiers temporaires sont dans `pdf_nova/tmp`.
- Taille max par fichier: 150 MB.
- OCR: installe `Tesseract OCR` sur Windows et ajoute `tesseract.exe` au `PATH`.
- Video MP4: `yt-dlp` installe via `requirements.txt`. `ffmpeg` recommande pour fusion audio+video.
- Office -> PDF: LibreOffice requis (`soffice`).
- Respecte les droits d'auteur et les conditions de la plateforme avant extraction.
