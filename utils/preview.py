"""
preview.py
Konversi bytes dokumen .docx → HTML untuk preview di browser.
Tidak ada file yang ditulis ke disk.
"""

import io
import base64
import mammoth


def docx_to_html(docx_bytes: bytes) -> str:
    """
    Konversi docx (bytes) menjadi HTML string menggunakan mammoth.
    Gambar dikonversi ke base64 data URI agar tidak butuh hosting file.
    """
    def _convert_image(image):
        with image.open() as img_file:
            img_bytes = img_file.read()
        b64 = base64.b64encode(img_bytes).decode("utf-8")
        content_type = image.content_type or "image/jpeg"
        return {"src": f"data:{content_type};base64,{b64}"}

    buf = io.BytesIO(docx_bytes)

    result = mammoth.convert_to_html(
        buf,
        convert_image=mammoth.images.img_element(_convert_image),
    )

    html_body = result.value

    # Wrap dengan styling agar tampilannya menyerupai dokumen Word A4
    html = f"""
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        body {{
            background: #e0e0e0;
            margin: 0;
            padding: 24px 0;
            font-family: 'Times New Roman', serif;
            font-size: 11pt;
        }}
        .page {{
            background: white;
            width: 21cm;
            min-height: 29.7cm;
            margin: 0 auto 24px auto;
            padding: 2cm;
            box-shadow: 0 2px 8px rgba(0,0,0,0.18);
            box-sizing: border-box;
        }}
        h1, h2, h3 {{
            font-family: 'Times New Roman', serif;
            font-size: 12pt;
            text-align: center;
            margin: 0 0 12px 0;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 8px;
        }}
        td, th {{
            font-size: 11pt;
            font-family: 'Times New Roman', serif;
            vertical-align: top;
            padding: 2px 4px;
        }}
        img {{
            max-width: 100%;
            height: auto;
            display: block;
            margin: 4px auto;
        }}
        p {{
            margin: 2px 0;
        }}
    </style>
    </head>
    <body>
        <div class="page">
            {html_body}
        </div>
    </body>
    </html>
    """
    return html
