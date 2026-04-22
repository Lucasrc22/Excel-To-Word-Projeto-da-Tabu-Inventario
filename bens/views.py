from django.shortcuts import render
from django.contrib import messages
from django.http import FileResponse, HttpResponse
from django.conf import settings
from .forms import UploadForm
from .services import salvar_imagens, importar_excel, listar_documentos_gerados
import os
import base64

def upload_view(request):
    download_file = None
    form = UploadForm()
    available_documents = listar_documentos_gerados()
    
    if request.method == "POST":
        form = UploadForm(request.POST, request.FILES)
        if form.is_valid():
            try:
                excel_file = request.FILES["excel"]
                imagens = request.FILES.getlist("imagens")

                # Ensure directories exist
                excel_dir = os.path.join(settings.MEDIA_ROOT, "uploads", "excels")
                os.makedirs(excel_dir, exist_ok=True)
                
                excel_path = os.path.join(excel_dir, excel_file.name)

                with open(excel_path, "wb+") as f:
                    for chunk in excel_file.chunks():
                        f.write(chunk)

                imagens_dict = salvar_imagens(imagens)
                doc_bytes, filename = importar_excel(excel_path, imagens_dict)
                messages.success(request, "Arquivo importado com sucesso!")
                
                # Store the filename in session to highlight the latest download
                request.session['latest_download'] = filename
                download_file = filename
                
                # Refresh the list of available documents
                available_documents = listar_documentos_gerados()
            except ValueError as e:
                messages.error(request, f"Erro na importação: {str(e)}")
            except Exception as e:
                messages.error(request, f"Erro inesperado: {str(e)}")
    
    return render(request, "bens/upload.html", {
        "form": form, 
        "download_file": download_file,
        "available_documents": available_documents
    })


def download_docx(request):
    """Download the generated DOCX file"""
    try:
        # Get filename from request (either from URL parameter or session)
        filename = request.GET.get('file')
        
        if not filename:
            # Try to get from session (latest download)
            filename = request.session.pop('latest_download', None)
        
        if not filename:
            messages.error(request, "Nenhum arquivo para baixar. Por favor, processe um inventário primeiro.")
            return render(request, "bens/upload.html", {
                "form": UploadForm(),
                "available_documents": listar_documentos_gerados()
            })
        
        # Construct full path - only allow files from outputs directory (security)
        if '/' in filename or '\\' in filename or '..' in filename:
            messages.error(request, "Caminho de arquivo inválido.")
            return render(request, "bens/upload.html", {
                "form": UploadForm(),
                "available_documents": listar_documentos_gerados()
            })
        
        file_path = os.path.join(settings.MEDIA_ROOT, 'outputs', filename)
        
        # Verify file exists
        if not os.path.exists(file_path) or not file_path.endswith('.docx'):
            messages.error(request, "Arquivo não encontrado.")
            return render(request, "bens/upload.html", {
                "form": UploadForm(),
                "available_documents": listar_documentos_gerados()
            })
        
        # Serve the file
        response = FileResponse(open(file_path, 'rb'), as_attachment=True)
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        response['Content-Type'] = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        
        return response
    except Exception as e:
        messages.error(request, f"Erro ao baixar arquivo: {str(e)}")
        return render(request, "bens/upload.html", {
            "form": UploadForm(),
            "available_documents": listar_documentos_gerados()
        })