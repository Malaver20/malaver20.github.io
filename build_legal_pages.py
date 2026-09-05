import sys
import os
import re
import docx

sys.stdout.reconfigure(encoding='utf-8')

HTML_HEAD = """<!DOCTYPE html>
<html lang="es" class="scroll-smooth">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>__TITLE__ | FacturaNex SAS</title>
    <link rel="icon" type="image/png" href="img/favicon.png" />
    
    <!-- SEO Meta Tags -->
    <meta name="description" content="__DESCRIPTION__">
    <meta name="keywords" content="FacturaNex, DIAN, Habeas Data, Legal, Colombia, Términos, Privacidad">

    <!-- Fonts -->
    <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">
    <link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@24,400,0,0" rel="stylesheet" />

    <!-- Tailwind CSS (Estilos Compilados Offline) -->
    <link rel="stylesheet" href="css/styles.css">
    
    <!-- AOS Animation Library -->
    <link href="https://unpkg.com/aos@2.3.1/dist/aos.css" rel="stylesheet">
    <script src="https://unpkg.com/aos@2.3.1/dist/aos.js"></script>

    <style>
        img {
            max-width: 100%;
        }
        nav img {
            max-height: 40px !important;
            width: auto !important;
            object-fit: contain !important;
        }
        body {
            background-color: #030712;
            color: #F8FAFC;
            overflow-x: hidden;
        }

        .glass {
            background: rgba(10, 17, 40, 0.7);
            backdrop-filter: blur(16px);
            -webkit-backdrop-filter: blur(16px);
            border: 1px solid rgba(255, 255, 255, 0.08);
        }
        
        .glass-card {
            background: linear-gradient(145deg, rgba(30, 41, 59, 0.4) 0%, rgba(15, 23, 42, 0.2) 100%);
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255, 255, 255, 0.08);
        }

        .text-gradient {
            background: linear-gradient(to right, #E91E63, #8B5CF6);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        .custom-scrollbar::-webkit-scrollbar {
            width: 6px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
            background: rgba(255, 255, 255, 0.02);
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
            background: rgba(216, 27, 96, 0.3);
            border-radius: 4px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
            background: rgba(216, 27, 96, 0.6);
        }
    </style>
</head>
<body class="antialiased selection:bg-primary selection:text-white relative">

    <!-- Fondo decorativo -->
    <div class="fixed inset-0 w-full h-full pointer-events-none -z-10 overflow-hidden">
        <div class="absolute top-0 left-1/4 w-96 h-96 bg-primary/10 rounded-full mix-blend-screen filter blur-[100px] opacity-40"></div>
        <div class="absolute top-1/3 right-1/4 w-96 h-96 bg-purple-600/10 rounded-full mix-blend-screen filter blur-[100px] opacity-30"></div>
    </div>

    <!-- Navegación -->
    <nav class="fixed w-full z-50 glass border-b border-white/5 py-4" id="navbar">
        <div class="max-w-7xl mx-auto px-6 lg:px-8 flex items-center justify-between">
            <a href="index.html" class="flex items-center gap-3">
                <img src="img/logo_facturanex_v3.png" alt="FacturaNex Logo" class="h-10 w-auto object-contain rounded-lg" style="max-height: 40px; max-width: 200px; height: 40px; width: auto; object-fit: contain;">
            </a>
            
            <div class="flex items-center gap-4">
                <a href="index.html" class="text-sm font-medium text-textMuted hover:text-white transition flex items-center gap-1">
                    <span class="material-symbols-outlined text-base">arrow_back</span>
                    <span>Volver al Inicio</span>
                </a>
                <a href="__PDF_FILE__" download class="hidden sm:inline-flex items-center gap-2 bg-primary hover:bg-primaryHover text-white text-xs font-bold py-2.5 px-4 rounded-xl shadow-lg transition">
                    <span class="material-symbols-outlined text-sm">download</span>
                    <span>Descargar PDF</span>
                </a>
            </div>
        </div>
    </nav>

    <!-- Header Principal -->
    <section class="pt-32 pb-12 relative border-b border-white/5 bg-surface/50">
        <div class="max-w-7xl mx-auto px-6 lg:px-8">
            <div class="max-w-3xl">
                <div class="inline-flex items-center gap-2 bg-primary/10 border border-primary/20 text-primary text-xs font-bold px-3 py-1 rounded-full uppercase tracking-wider mb-4">
                    <span class="material-symbols-outlined text-sm">gavel</span>
                    <span>Documento Oficial</span>
                </div>
                <h1 class="text-3xl md:text-5xl font-extrabold text-white mb-4 leading-tight">
                    __HEADING_TITLE__
                </h1>
                <p class="text-textMuted text-sm md:text-base leading-relaxed mb-6">
                    FacturaNex SAS — NIT 902.049.529-8 | Última actualización: 5 de septiembre de 2026
                </p>
                <div class="flex flex-wrap gap-3">
                    <a href="__PDF_FILE__" download class="inline-flex items-center gap-2 bg-primary hover:bg-primaryHover text-white text-sm font-bold py-3 px-6 rounded-xl shadow-lg transition">
                        <span class="material-symbols-outlined text-base">picture_as_pdf</span>
                        <span>Descargar Versión PDF</span>
                    </a>
                    <a href="https://wa.me/573155155512?text=Hola,%20tengo%20una%20consulta%20legal%20o%20de%20privacidad" target="_blank" class="inline-flex items-center gap-2 bg-white/5 hover:bg-white/10 text-white text-sm font-bold py-3 px-6 rounded-xl border border-white/10 transition">
                        <span class="material-symbols-outlined text-base">chat</span>
                        <span>Consultar por WhatsApp</span>
                    </a>
                </div>
            </div>
        </div>
    </section>

    <!-- Contenido Principal con Sidebar -->
    <div class="max-w-7xl mx-auto px-6 lg:px-8 py-12">
        <div class="grid lg:grid-cols-12 gap-10">

            <!-- Sidebar con Índice (Desktop) -->
            <aside class="lg:col-span-4 hidden lg:block">
                <div class="sticky top-28 glass p-6 rounded-2xl max-h-[calc(100vh-140px)] overflow-y-auto custom-scrollbar">
                    <h3 class="text-xs font-bold text-primary uppercase tracking-widest mb-4 flex items-center gap-2">
                        <span class="material-symbols-outlined text-sm">toc</span>
                        Índice de Contenido
                    </h3>
                    <nav class="space-y-1 text-xs">
                        __SIDEBAR_LINKS__
                    </nav>
                </div>
            </aside>

            <!-- Cuerpo del Documento -->
            <main class="lg:col-span-8 space-y-8">
                __DOCUMENT_BODY__
            </main>
        </div>
    </div>
"""

HTML_FOOTER = """
    <!-- Footer -->
    <footer class="border-t border-white/5 pt-12 pb-8 bg-surface mt-16">
        <div class="max-w-7xl mx-auto px-6 flex flex-col md:flex-row justify-between items-center gap-6">
            <div class="flex items-center gap-4">
                <img src="img/logo_facturanex_v3.png" alt="FacturaNex Logo" class="h-10 w-auto object-contain rounded" style="max-height: 40px; max-width: 180px; height: 40px; width: auto; object-fit: contain;">
                <div class="text-left">
                    <p class="text-white font-bold tracking-widest text-sm">FacturaNex SAS</p>
                    <p class="text-textMuted text-xs mt-1">Conexión Directa con la DIAN — NIT 902.049.529-8</p>
                </div>
            </div>
            
            <div class="flex flex-wrap justify-center gap-6 text-xs text-textMuted">
                <a href="politica-de-privacidad.html" class="hover:text-primary transition">Política de Privacidad</a>
                <a href="terminos-y-condiciones.html" class="hover:text-primary transition">Términos y Condiciones</a>
                <a href="habeas-data.html" class="hover:text-primary transition">Habeas Data</a>
                <a href="index.html" class="hover:text-white transition">Inicio</a>
            </div>
        </div>
        <div class="max-w-7xl mx-auto px-6 mt-8 pt-6 border-t border-white/5 text-center text-xs text-textMuted opacity-60">
            © 2026 FacturaNex SAS. Impulsando y digitalizando el crecimiento de las empresas en Colombia.
        </div>
    </footer>

    <!-- Initialize AOS -->
    <script>
        AOS.init({ once: true, offset: 50, duration: 600 });
    </script>
</body>
</html>
"""

def parse_docx_to_html(docx_filename, title, pdf_file, is_habeas=False):
    doc = docx.Document(docx_filename)
    
    sections = []
    current_sec = None
    
    sidebar_links = []
    
    for p in doc.paragraphs:
        txt = p.text.strip()
        if not txt:
            continue
        
        if re.match(r'^\d+(\.\d+)*\.\s+', txt) or (txt.isupper() and len(txt) < 80 and not txt.startswith("FACTURANEX")):
            sec_id = "section-" + re.sub(r'[^a-zA-Z0-9]', '-', txt.split()[0]).lower() if re.match(r'^\d+', txt) else f"section-{len(sections)+1}"
            
            if current_sec:
                sections.append(current_sec)
            
            current_sec = {
                "id": sec_id,
                "title": txt,
                "paragraphs": []
            }
            
            clean_title = txt if len(txt) < 55 else txt[:52] + "..."
            sidebar_links.append(f'<a href="#{sec_id}" class="block py-1.5 px-3 rounded-lg text-textMuted hover:text-white hover:bg-white/5 transition border-l-2 border-transparent hover:border-primary truncate">{clean_title}</a>')
        else:
            if current_sec is None:
                current_sec = {
                    "id": "section-intro",
                    "title": "Introducción",
                    "paragraphs": []
                }
                sidebar_links.append(f'<a href="#section-intro" class="block py-1.5 px-3 rounded-lg text-textMuted hover:text-white hover:bg-white/5 transition border-l-2 border-transparent hover:border-primary truncate">Introducción</a>')
            
            current_sec["paragraphs"].append(txt)
    
    if current_sec:
        sections.append(current_sec)

    body_html_parts = []
    
    for sec in sections:
        sec_html = f'<div id="{sec["id"]}" class="glass-card p-6 md:p-8 rounded-2xl border border-white/5 scroll-mt-28" data-aos="fade-up">\n'
        sec_html += f'  <h2 class="text-xl font-bold text-white mb-4 flex items-center gap-2 border-b border-white/5 pb-3"><span class="w-2 h-2 bg-primary rounded-full"></span> {sec["title"]}</h2>\n'
        sec_html += '  <div class="space-y-3 text-sm text-textMuted leading-relaxed">\n'
        
        for p_text in sec["paragraphs"]:
            if p_text.startswith("•") or p_text.startswith("-"):
                sec_html += f'    <p class="flex items-start gap-2 pl-2"><span class="text-primary font-bold">✓</span> <span>{p_text[1:].strip()}</span></p>\n'
            elif p_text.startswith("http"):
                sec_html += f'    <p><a href="{p_text}" target="_blank" class="text-primary hover:underline font-medium">{p_text}</a></p>\n'
            else:
                sec_html += f'    <p>{p_text}</p>\n'
                
        sec_html += '  </div>\n</div>\n'
        body_html_parts.append(sec_html)

    if is_habeas:
        form_html = """
        <div id="formulario-habeas" class="glass-card p-8 rounded-2xl border border-primary/30 shadow-[0_0_30px_rgba(216,27,96,0.1)] scroll-mt-28" data-aos="fade-up">
            <div class="flex items-center gap-3 mb-4">
                <div class="h-10 w-10 rounded-xl bg-primary/20 flex items-center justify-center text-primary font-bold">
                    <span class="material-symbols-outlined">edit_document</span>
                </div>
                <div>
                    <h3 class="text-xl font-bold text-white">Solicitud Digital de Habeas Data</h3>
                    <p class="text-xs text-textMuted">Ejerce tus derechos de Consulta, Actualización o Supresión de Datos Personales</p>
                </div>
            </div>
            
            <form onsubmit="handleHabeasSubmit(event)" class="space-y-4 text-xs">
                <div class="grid md:grid-cols-2 gap-4">
                    <div>
                        <label class="block text-white font-bold mb-1">Nombre Completo o Razón Social *</label>
                        <input type="text" id="hd_nombre" required class="w-full bg-surface border border-white/10 rounded-xl px-4 py-3 text-white focus:border-primary outline-none" placeholder="Ej: Maria Perez">
                    </div>
                    <div>
                        <label class="block text-white font-bold mb-1">Cédula / NIT *</label>
                        <input type="text" id="hd_identificacion" required class="w-full bg-surface border border-white/10 rounded-xl px-4 py-3 text-white focus:border-primary outline-none" placeholder="Ej: 1.098.765.432">
                    </div>
                </div>

                <div class="grid md:grid-cols-2 gap-4">
                    <div>
                        <label class="block text-white font-bold mb-1">Correo Electrónico *</label>
                        <input type="email" id="hd_email" required class="w-full bg-surface border border-white/10 rounded-xl px-4 py-3 text-white focus:border-primary outline-none" placeholder="contacto@empresa.com">
                    </div>
                    <div>
                        <label class="block text-white font-bold mb-1">Teléfono de Contacto *</label>
                        <input type="tel" id="hd_telefono" required class="w-full bg-surface border border-white/10 rounded-xl px-4 py-3 text-white focus:border-primary outline-none" placeholder="300 123 4567">
                    </div>
                </div>

                <div>
                    <label class="block text-white font-bold mb-1">Tipo de Solicitud *</label>
                    <select id="hd_tipo" required class="w-full bg-surface border border-white/10 rounded-xl px-4 py-3 text-white focus:border-primary outline-none">
                        <option value="Consulta de Datos">Consulta de Información Almacenada</option>
                        <option value="Actualización / Rectificación">Actualización o Rectificación de Datos</option>
                        <option value="Supresión / Eliminación">Supresión / Eliminación de Datos (Cancelación)</option>
                        <option value="Revocatoria de Autorización">Revocatoria de Autorización</option>
                    </select>
                </div>

                <div>
                    <label class="block text-white font-bold mb-1">Descripción de la Solicitud *</label>
                    <textarea id="hd_mensaje" rows="4" required class="w-full bg-surface border border-white/10 rounded-xl px-4 py-3 text-white focus:border-primary outline-none" placeholder="Describe claramente el motivo de tu solicitud..."></textarea>
                </div>

                <button type="submit" class="w-full py-4 bg-primary hover:bg-primaryHover text-white font-bold rounded-xl transition shadow-lg flex items-center justify-center gap-2">
                    <span class="material-symbols-outlined text-base">send</span>
                    <span>Enviar Solicitud a FacturaNex</span>
                </button>
            </form>

            <script>
                function handleHabeasSubmit(e) {
                    e.preventDefault();
                    const nombre = document.getElementById('hd_nombre').value;
                    const id = document.getElementById('hd_identificacion').value;
                    const email = document.getElementById('hd_email').value;
                    const tipo = document.getElementById('hd_tipo').value;
                    const msg = document.getElementById('hd_mensaje').value;

                    const text = `Solicitud de Habeas Data FacturaNex%0A- Nombre: ${nombre}%0A- ID/NIT: ${id}%0A- Correo: ${email}%0A- Tipo: ${tipo}%0A- Detalle: ${msg}`;
                    window.open(`https://wa.me/573155155512?text=${text}`, '_blank');
                }
            </script>
        </div>
        """
        body_html_parts.insert(0, form_html)

    sidebar_html = "\n".join(sidebar_links)
    body_html = "\n".join(body_html_parts)

    head_part = HTML_HEAD.replace("__TITLE__", title) \
                         .replace("__DESCRIPTION__", f"{title} de FacturaNex SAS. Transparencia, legalidad y protección de datos.") \
                         .replace("__PDF_FILE__", pdf_file) \
                         .replace("__HEADING_TITLE__", title) \
                         .replace("__SIDEBAR_LINKS__", sidebar_html) \
                         .replace("__DOCUMENT_BODY__", body_html)

    return head_part + HTML_FOOTER

if __name__ == "__main__":
    # 1. Política de Privacidad
    html_privacidad = parse_docx_to_html(
        "POLÍTICA DE PRIVACIDAD Y TRATAMIENTO DE DATOS PERSONALES.docx",
        "Política de Privacidad y Tratamiento de Datos",
        "politica_de_privacidad_facturanex.pdf"
    )
    with open("politica-de-privacidad.html", "w", encoding="utf-8") as f:
        f.write(html_privacidad)
    print("Generated politica-de-privacidad.html")

    # 2. Términos y Condiciones
    html_terminos = parse_docx_to_html(
        "TÉRMINOS Y CONDICIONES DE USO Y PRESTACIÓN DEL SERVICIO.docx",
        "Términos y Condiciones del Servicio",
        "terminos_y_condiciones_facturanex.pdf"
    )
    with open("terminos-y-condiciones.html", "w", encoding="utf-8") as f:
        f.write(html_terminos)
    print("Generated terminos-y-condiciones.html")

    # 3. Habeas Data
    html_habeas = parse_docx_to_html(
        "HABEAS DATA.docx",
        "Procedimiento de Habeas Data",
        "habeas_data_facturanex.pdf",
        is_habeas=True
    )
    with open("habeas-data.html", "w", encoding="utf-8") as f:
        f.write(html_habeas)
    print("Generated habeas-data.html")
