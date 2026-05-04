# models.py
from django.db import models

class Bem(models.Model):
    numero_bem   = models.CharField(max_length=50, unique=True)
    area         = models.CharField(max_length=255, blank=True, null=True)
    localizacao  = models.CharField(max_length=255, blank=True, null=True)
    centro_custo = models.CharField(max_length=100, blank=True, null=True)
    descricao_cc = models.CharField(max_length=255, blank=True, null=True)
    responsavel  = models.CharField(max_length=255, blank=True, null=True)
    descricao_bem = models.CharField(max_length=255, blank=True, null=True)
    marca        = models.CharField(max_length=100, blank=True, null=True)
    modelo       = models.CharField(max_length=100, blank=True, null=True)
    narrativa    = models.TextField(blank=True, null=True)
    estado       = models.CharField(max_length=100, blank=True, null=True)

    imagem_1 = models.ImageField(upload_to="uploads/imagens/", blank=True, null=True)
    imagem_2 = models.ImageField(upload_to="uploads/imagens/", blank=True, null=True)
    imagem_3 = models.ImageField(upload_to="uploads/imagens/", blank=True, null=True) 
    imagem_4 = models.ImageField(upload_to="uploads/imagens/", blank=True, null=True) 
    imagem_5 = models.ImageField(upload_to="uploads/imagens/", blank=True, null=True)  
    imagem_6 = models.ImageField(upload_to="uploads/imagens/", blank=True, null=True)  

    criado_em    = models.DateTimeField(auto_now_add=True)
    atualizado_em = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Bem"
        verbose_name_plural = "Bens"
        ordering = ["numero_bem"]

    def __str__(self) -> str:
        return f"Bem #{self.numero_bem}"