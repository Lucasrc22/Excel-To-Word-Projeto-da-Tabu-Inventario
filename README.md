Conversor de Inventário Excel → Word
Visão geral

Este projeto é uma aplicação baseada em Django para processamento de arquivos de inventário em Excel e geração automática de relatórios em Word (.docx).

O sistema realiza leitura estruturada, validação rígida de schema e exportação determinística dos dados, mantendo fidelidade total ao conteúdo original da planilha.

Tecnologias utilizadas
Python 3.10+
Django
pandas
openpyxl
python-docx
venv (ambiente virtual)

1. Clonar repositório
git clone <URL_DO_REPOSITORIO>
cd inventory_project

2. Criar ambiente virtual
python -m venv venv

3. Ativar ambiente virtual

Windows

venv\Scripts\activate

Linux / macOS

source venv/bin/activate

4. Instalar dependências

O projeto utiliza requirements.txt como fonte oficial de dependências.

pip install --upgrade pip
pip install -r requirements.txt

Configuração do banco de dados
python manage.py makemigrations
python manage.py migrate

Execução
python manage.py runserver

Acesso local:

http://127.0.0.1:8000/

Evoluções futuras
API REST (Django REST Framework)
Celery + Redis para filas
Exportação PDF
Interface administrativa de upload/download
Validação de schema configurável
