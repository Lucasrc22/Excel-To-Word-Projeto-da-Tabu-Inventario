import pandas as pd
import os
from django.conf import settings
from .models import Bem
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm


def salvar_imagens(files):
    """Save uploaded image files and return a dict with image paths"""
    from django.core.files.storage import default_storage
    from PIL import Image
    import io
    
    paths = {}
    base_path = "uploads/imagens"

    for f in files:
        try:
            # Validate and convert image before saving
            img = Image.open(f)
            
            # Convert to RGB if needed (handles PNG with transparency, etc.)
            if img.mode in ('RGBA', 'LA', 'P'):
                # Create a white background
                background = Image.new('RGB', img.size, (255, 255, 255))
                background.paste(img, mask=img.split()[-1] if img.mode in ('RGBA', 'LA') else None)
                img = background
            
            # Save as JPEG to ensure compatibility
            img_io = io.BytesIO()
            img.save(img_io, format='JPEG', quality=85)
            img_io.seek(0)
            
            # Get filename and change extension to .jpg
            nome = os.path.splitext(f.name)[0]
            file_name = f"{nome}.jpg"
            
            # Save using Django's storage system
            file_path = os.path.join(base_path, file_name)
            saved_path = default_storage.save(file_path, img_io)
            
            # Store the saved path (relative to MEDIA_ROOT)
            paths[nome] = saved_path
            print(f"Imagem {f.name} salva com sucesso: {saved_path}")
        except Exception as e:
            print(f"Erro ao salvar imagem {f.name}: {str(e)}")
            # Continue with other files even if one fails
            continue

    return paths

def importar_excel(excel_path, imagens_dict):
    # Try to read the excel file with proper header handling
    df = pd.read_excel(excel_path, header=0)
    
    # Strip whitespace from column names
    df.columns = df.columns.str.strip()
    
    required_columns = [
        "Numero bem",
        "IMOBILIZADO POR AREA",
        "Localizacao",
        "Centro custo",
        "Descricao do CC",
        "Responsavel",
        "Descricao/TAG",
        "Marca",
        "Modelo",
        "Narrativa",
        "Estado Fisico",
    ]
    
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        available = list(df.columns)
        raise ValueError(
            f"Colunas ausentes no arquivo Excel: {missing_columns}\n"
            f"Colunas disponíveis: {available}"
        )

    for _, row in df.iterrows():
        numero = str(row["Numero bem"]).strip()

        img1_path = imagens_dict.get(numero)
        img2_path = imagens_dict.get(f"{numero}A")

        bem, _ = Bem.objects.update_or_create(
            numero_bem=numero,
            defaults={
                "area": row["IMOBILIZADO POR AREA"],
                "localizacao": row["Localizacao"],
                "centro_custo": row["Centro custo"],
                "descricao_cc": row["Descricao do CC"],
                "responsavel": row["Responsavel"],
                "tag": row["Descricao/TAG"],
                "marca": row["Marca"],
                "modelo": row["Modelo"],
                "narrativa": row["Narrativa"],
                "estado": row["Estado Fisico"],
            }
        )

        if img1_path:
            bem.imagem_1 = img1_path
        if img2_path:
            bem.imagem_2 = img2_path

        bem.save()
    
    # Generate unified document with all items
    result = gerar_docx_unificado()
    
    # Delete uploaded files after successful processing
    deletar_arquivos()
    
    return result

def gerar_docx_unificado():
    """Generates a single unified Word document with all registered assets and returns BytesIO"""
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    from io import BytesIO
    
    bens = Bem.objects.all().order_by('numero_bem')
    
    if not bens.exists():
        raise ValueError("Nenhum bem registrado para gerar documento")
    
    # Create a new document
    doc = Document()
    
    # Add title
    title = doc.add_heading('Inventário de Ativos', 0)
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    doc.add_paragraph()  # Add spacing
    
    # Add each bem to the document
    for bem in bens:
        # Add bem number as heading
        doc.add_heading(f'Bem #{bem.numero_bem}', level=1)
        
        # Add bem details in a table format for better organization
        table = doc.add_table(rows=11, cols=2)
        table.style = 'Light Grid Accent 1'
        
        data = [
            ('Número do Bem', bem.numero_bem),
            ('Descrição', bem.descricao_cc),
            ('Localização', bem.localizacao),
            ('Centro de Custo', bem.centro_custo),
            ('Responsável', bem.responsavel),
            ('TAG', bem.tag),
            ('Marca', bem.marca),
            ('Modelo', bem.modelo),
            ('Narrativa', bem.narrativa),
            ('Estado Físico', bem.estado),
            ('Área', bem.area),
        ]
        
        for i, (label, value) in enumerate(data):
            table.rows[i].cells[0].text = label
            table.rows[i].cells[1].text = str(value) if value else ''
        
        # Add images if available
        if bem.imagem_1 or bem.imagem_2:
            doc.add_heading('Imagens', level=2)
            
            if bem.imagem_1 and bem.imagem_1.name:
                try:
                    # Get the full path to the image file
                    image_path = bem.imagem_1.path
                    
                    # Validate file exists and has content
                    if not os.path.exists(image_path):
                        doc.add_paragraph(f'Erro: Imagem 1 não encontrada em {image_path}')
                    elif os.path.getsize(image_path) == 0:
                        doc.add_paragraph(f'Erro: Imagem 1 está vazia ({image_path})')
                    else:
                        # Verify image is valid before adding
                        try:
                            from PIL import Image as PILImage
                            test_img = PILImage.open(image_path)
                            test_img.verify()
                            
                            # If verification passed, add the picture
                            doc.add_paragraph('Imagem 1:')
                            doc.add_picture(image_path, width=Mm(80))
                        except Exception as verify_error:
                            doc.add_paragraph(f'Aviso: Imagem 1 pode estar corrompida. Pulando... ({type(verify_error).__name__})')
                except Exception as e:
                    doc.add_paragraph(f'Erro ao processar imagem 1: {type(e).__name__}: {str(e)}')
            
            if bem.imagem_2 and bem.imagem_2.name:
                try:
                    # Get the full path to the image file
                    image_path = bem.imagem_2.path
                    
                    # Validate file exists and has content
                    if not os.path.exists(image_path):
                        doc.add_paragraph(f'Erro: Imagem 2 não encontrada em {image_path}')
                    elif os.path.getsize(image_path) == 0:
                        doc.add_paragraph(f'Erro: Imagem 2 está vazia ({image_path})')
                    else:
                        # Verify image is valid before adding
                        try:
                            from PIL import Image as PILImage
                            test_img = PILImage.open(image_path)
                            test_img.verify()
                            
                            # If verification passed, add the picture
                            doc.add_paragraph('Imagem 2:')
                            doc.add_picture(image_path, width=Mm(80))
                        except Exception as verify_error:
                            doc.add_paragraph(f'Aviso: Imagem 2 pode estar corrompida. Pulando... ({type(verify_error).__name__})')
                except Exception as e:
                    doc.add_paragraph(f'Erro ao processar imagem 2: {type(e).__name__}: {str(e)}')
        
        # Add page break between items (except for the last one)
        if bem != bens.last():
            doc.add_page_break()
    
    # Save the document to BytesIO and also to disk with timestamp
    from datetime import datetime
    
    doc_bytes = BytesIO()
    doc.save(doc_bytes)
    doc_bytes.seek(0)
    
    # Also save to disk for persistence
    output_dir = os.path.join(settings.MEDIA_ROOT, 'outputs')
    os.makedirs(output_dir, exist_ok=True)
    
    # Create filename with timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_filename = f'Inventario_Unificado_{timestamp}.docx'
    output_path = os.path.join(output_dir, output_filename)
    
    with open(output_path, 'wb') as f:
        f.write(doc_bytes.getvalue())
    
    doc_bytes.seek(0)
    return doc_bytes, output_filename


def gerar_docx(bem):
    """Generates a single Word document for one specific asset (kept for compatibility)"""
    template_path = "templates_docx/template.docx"
    doc = DocxTemplate(template_path)

    context = {
        "numero_bem": bem.numero_bem,
        "descricao": bem.descricao_cc,
        "localizacao": bem.localizacao,
        "centro_custo": bem.centro_custo,
        "responsavel": bem.responsavel,
        "marca": bem.marca,
        "modelo": bem.modelo,
        "narrativa": bem.narrativa,
        "estado": bem.estado,
    }

    if bem.imagem_1:
        context["imagem_1"] = InlineImage(doc, bem.imagem_1.path, width=Mm(80))
    else:
        context["imagem_1"] = ""

    if bem.imagem_2:
        context["imagem_2"] = InlineImage(doc, bem.imagem_2.path, width=Mm(80))
    else:
        context["imagem_2"] = ""

    doc.render(context)

    output_path = f"media/output/{bem.numero_bem}.docx"
    doc.save(output_path)

    return output_path
#Delete files after processing in media/uploads (for testing purposes), only manteins the docx files in media/output
def deletar_arquivos():
    """Deletes all files in the media/uploads directory (for testing purposes)"""
    upload_dir = os.path.join(settings.MEDIA_ROOT, "uploads")
    for root, dirs, files in os.walk(upload_dir):
        for file in files:
            try:
                os.remove(os.path.join(root, file))
            except Exception as e:
                print(f"Erro ao deletar arquivo {file}: {str(e)}")

def listar_documentos_gerados():
    """List all generated DOCX documents"""
    output_dir = os.path.join(settings.MEDIA_ROOT, 'outputs')
    if not os.path.exists(output_dir):
        return []
    
    files = []
    for filename in sorted(os.listdir(output_dir), reverse=True):
        if filename.endswith('.docx'):
            filepath = os.path.join(output_dir, filename)
            file_size = os.path.getsize(filepath)
            file_time = os.path.getmtime(filepath)
            from datetime import datetime
            formatted_time = datetime.fromtimestamp(file_time).strftime('%d/%m/%Y %H:%M:%S')
            files.append({
                'filename': filename,
                'size': f"{file_size / (1024*1024):.2f}MB" if file_size > 1024*1024 else f"{file_size / 1024:.2f}KB",
                'date': formatted_time,
                'url_name': filename
            })
    return files