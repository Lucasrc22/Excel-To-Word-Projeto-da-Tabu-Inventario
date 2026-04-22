from django.db import models

class Bem(models.Model):
    numero_bem = models.CharField(max_length=50, unique=True)

    area = models.CharField(max_length=200)
    localizacao = models.CharField(max_length=200)
    centro_custo = models.CharField(max_length=200)
    descricao_cc = models.TextField()

    responsavel = models.CharField(max_length=200)
    tag = models.CharField(max_length=200)

    marca = models.CharField(max_length=200)
    modelo = models.CharField(max_length=200)

    narrativa = models.TextField()
    estado = models.CharField(max_length=50)

    imagem_1 = models.ImageField(upload_to="uploads/imagens/", null=True, blank=True)
    imagem_2 = models.ImageField(upload_to="uploads/imagens/", null=True, blank=True)

    criado_em = models.DateTimeField(auto_now_add=True)
