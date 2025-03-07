from setuptools import setup, find_packages
import os

# Crear directorios necesarios
dirs = [
    'core',
    'generators',
    'ui',
    'utils',
    'templates'
]

for directory in dirs:
    os.makedirs(directory, exist_ok=True)

# Asegurarse de que existan los archivos __init__.py en cada módulo
for directory in dirs:
    init_file = os.path.join(directory, '__init__.py')
    if not os.path.exists(init_file):
        with open(init_file, 'w') as f:
            f.write('# Inicializador del módulo\n')

setup(
    name="AutomatizacionPartidas",
    version="1.0",
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        "python-docx>=0.8.11",
        "docx2pdf>=0.1.8",
        "pandas>=1.3.0",
        "openpyxl>=3.0.7",
        "fpdf>=1.7.2",
        "babel>=2.9.1",
        "selenium>=4.0.0",
        "webdriver-manager>=3.5.2",
        "tkcalendar>=1.6.1",
    ],
    author="Tu Nombre",
    author_email="tu@email.com",
    description="Sistema de automatización de documentos por partidas",
    keywords="automatización, documentos, partidas, facturas, xml",
    url="https://github.com/yourusername/proyecto",
    classifiers=[
        "Development Status :: 4 - Beta",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
    ],
    entry_points={
        'console_scripts': [
            'app=main:run',
        ],
    },
)
