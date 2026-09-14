"""Helper comun pentru inserarea logo-ului oficial in documentele generate
(Fisa limita, RetDozare), atat in format .xlsx cat si .pdf.

Fisierul de logo e cautat la config.CALE_LOGO (implicit
dozare_titan/assets/logo.png — acolo se afla momentan un PLACEHOLDER, vezi
dozare_titan/assets/README.md). Daca fisierul lipseste sau nu poate fi
incarcat (ex. Pillow neinstalat), functiile de mai jos nu fac nimic /
returneaza None, iar codul apelant revine la varianta text "ZIROM TITANIUM"
ca rezerva — nu se intrerupe generarea documentului.
"""

import os

from ..config import CALE_LOGO


def _raport_aspect(cale):
    """Inaltime/latime a imaginii, folosit ca sa pastram proportiile la
    redimensionare. Necesita Pillow — daca nu e instalat, presupunem un
    raport tipic de logo lat (3.5:1)."""
    try:
        from PIL import Image as ImaginePIL
        with ImaginePIL.open(cale) as im:
            return im.height / im.width
    except Exception:
        return 1 / 3.5


def adauga_logo_xlsx(ws, celula="A1", latime_px=120):
    """Insereaza logo-ul in foaia xlsx, ancorat la celula data, pastrand
    raportul de aspect. Returneaza True daca a reusit, False daca fisierul
    de logo lipseste sau imaginea n-a putut fi incarcata (caz in care
    apelantul trebuie sa scrie textul "ZIROM TITANIUM" ca rezerva)."""
    if not os.path.exists(CALE_LOGO):
        return False
    try:
        from openpyxl.drawing.image import Image as ImagineXlsx
        img = ImagineXlsx(CALE_LOGO)
        img.width = latime_px
        img.height = int(latime_px * _raport_aspect(CALE_LOGO))
        ws.add_image(img, celula)
        return True
    except Exception:
        return False


def logo_flowable_pdf(latime_cm=3.6):
    """Returneaza un flowable reportlab (Image) cu logo-ul, sau None daca
    fisierul lipseste ori nu poate fi incarcat — apelantul foloseste un
    Paragraph cu textul "ZIROM TITANIUM" ca rezerva in acest caz."""
    if not os.path.exists(CALE_LOGO):
        return None
    try:
        from reportlab.platypus import Image as ImaginePdf
        from reportlab.lib.units import cm
        latime = latime_cm * cm
        return ImaginePdf(CALE_LOGO, width=latime, height=latime * _raport_aspect(CALE_LOGO))
    except Exception:
        return None
