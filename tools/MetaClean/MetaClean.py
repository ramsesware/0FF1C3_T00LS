# GNU GENERAL PUBLIC LICENSE
# Version 3, 29 June 2007

# Copyright (C) 2024 Moisés Ceñera Fernández

# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU General Public License for more details.

# You should have received a copy of the GNU General Public License
# along with this program. If not, see <https://www.gnu.org/licenses/>.

import wx
import os
from PyPDF2 import PdfReader, PdfWriter
from docx import Document
from openpyxl import load_workbook
from pptx import Presentation
from PIL import Image
import zipfile
import shutil
from datetime import datetime
import xml.etree.ElementTree as ET
from mutagen import File as MutagenFile
from hachoir.parser import createParser
from hachoir.metadata import extractMetadata
from hachoir.stream import FileOutputStream
from hachoir.editor import createEditor


class MainApp(wx.App):
    def OnInit(self):
        self.frame = MetadataAnalyzerFrame(None, title="MetaClean", size=(800, 600))
        self.frame.Show()
        return True

class MetadataAnalyzerFrame(wx.Frame):
    def __init__(self, *args, **kw):
        super(MetadataAnalyzerFrame, self).__init__(*args, **kw)
        
        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        instructions = wx.StaticText(panel, label="Seleccione una acción para analizar o limpiar los metadatos de archivos.")
        instruction_font = wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
        instructions.SetFont(instruction_font)
        main_sizer.Add(instructions, 0, wx.ALL | wx.CENTER, 10)
        
        buttons_sizer = wx.GridBagSizer(10, 10) 
        
        button_font = wx.Font(11, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        button_size = (250, 50)
        
        select_file_btn = wx.Button(panel, label="📄 Seleccionar Archivo", size=button_size)
        select_file_btn.SetFont(button_font)
        select_file_btn.Bind(wx.EVT_BUTTON, self.on_select_file)
        buttons_sizer.Add(select_file_btn, pos=(0, 0), flag=wx.EXPAND | wx.ALL, border=5)
        
        remove_file_btn = wx.Button(panel, label="🗑️ Eliminar Metadatos (Archivo)", size=button_size)
        remove_file_btn.SetFont(button_font)
        remove_file_btn.Bind(wx.EVT_BUTTON, self.on_remove_metadata_file)
        buttons_sizer.Add(remove_file_btn, pos=(1, 0), flag=wx.EXPAND | wx.ALL, border=5)
        
        select_directory_btn = wx.Button(panel, label="📂 Seleccionar Carpeta", size=button_size)
        select_directory_btn.SetFont(button_font)
        select_directory_btn.Bind(wx.EVT_BUTTON, self.on_select_directory)
        buttons_sizer.Add(select_directory_btn, pos=(0, 1), flag=wx.EXPAND | wx.ALL, border=5)
        
        remove_directory_btn = wx.Button(panel, label="🗑️📂 Eliminar Metadatos (Carpeta)", size=button_size)
        remove_directory_btn.SetFont(button_font)
        remove_directory_btn.Bind(wx.EVT_BUTTON, self.on_remove_metadata_directory)
        buttons_sizer.Add(remove_directory_btn, pos=(1, 1), flag=wx.EXPAND | wx.ALL, border=5)
        
        clear_results_btn = wx.Button(panel, label="🧹 Limpiar Resultados", size=button_size)
        clear_results_btn.SetFont(button_font)
        clear_results_btn.Bind(wx.EVT_BUTTON, self.on_clear_results)
        buttons_sizer.Add(clear_results_btn, pos=(2, 0), span=(1,2), flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, border=5)
        
        buttons_sizer.AddGrowableCol(0, 1)
        buttons_sizer.AddGrowableCol(1, 1)
        
        main_sizer.Add(buttons_sizer, 0, wx.ALL | wx.CENTER, 15)
        
        self.result_text_metadata = wx.TextCtrl(panel, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL, size=(750, 250))
        main_sizer.Add(self.result_text_metadata, 1, wx.ALL | wx.EXPAND, 10)
        
        panel.SetSizer(main_sizer)

    def on_select_file(self, event):
        with wx.FileDialog(self, "Seleccione un archivo", style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as file_dialog:
            if file_dialog.ShowModal() == wx.ID_CANCEL:
                return
            path = file_dialog.GetPath()
            metadata = analyze_metadata(self, path)
            self.display_metadata(metadata)

    def on_select_directory(self, event):
        with wx.DirDialog(self, "Seleccione una carpeta", style=wx.DD_DEFAULT_STYLE) as dir_dialog:
            if dir_dialog.ShowModal() == wx.ID_CANCEL:
                return None
            directory_path = dir_dialog.GetPath()
            if not directory_path:
                return None  
            file_list = [os.path.join(directory_path, f) for f in os.listdir(directory_path) if os.path.isfile(os.path.join(directory_path, f))]
            self.display_directory_metadata(analyze_metadata_directory(self, file_list))
            
    def on_remove_metadata_file(self, event):
        with wx.FileDialog(self, "Seleccione un archivo", style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as file_dialog:
            if file_dialog.ShowModal() == wx.ID_CANCEL:
                return
            path = file_dialog.GetPath()
            result = remove_metadata_file(self, path)
            self.display_result(result)

    def on_remove_metadata_directory(self, event):
        with wx.DirDialog(self, "Seleccione una carpeta") as dir_dialog:
            if dir_dialog.ShowModal() == wx.ID_CANCEL:
                return
            directory_path = dir_dialog.GetPath()
            if not directory_path:
                return None  
            file_list = [os.path.join(directory_path, f) for f in os.listdir(directory_path) if os.path.isfile(os.path.join(directory_path, f))]
            results = remove_metadata_directory(self, file_list)
            self.display_directory_results(results)

    def on_clear_results(self, event):
        self.result_text_metadata.Clear()

    def display_metadata(self, data):
        self.result_text_metadata.Clear()
        if data:
            for key, value in data.items():
                self.result_text_metadata.AppendText(f"{key}: {value}\n")
        else:
            self.result_text_metadata.AppendText("Advertencia: No se encontraron metadatos o el archivo seleccionado es inválido.")

    def display_result(self, data):
        self.result_text_metadata.Clear()
        if data:
            self.result_text_metadata.AppendText(f"{data}\n")
        else:
            self.result_text_metadata.AppendText("Advertencia: No se pudo eliminar los metadatos o el archivo seleccionado es inválido.")

    def display_directory_metadata(self, data):
        self.result_text_metadata.Clear()
        if data:
            for file_data in data:
                filename = file_data.get("filename", "Archivo desconocido")
                metadata = file_data.get("metadata", {})
                self.result_text_metadata.AppendText(f"Archivo: {filename}\n")
                if isinstance(metadata, dict):
                    for key, value in metadata.items():
                        self.result_text_metadata.AppendText(f"  {key}: {value}\n")
                else:
                    self.result_text_metadata.AppendText(f"  {metadata}\n")
                self.result_text_metadata.AppendText("\n" + "-" * 40 + "\n\n")
        else:
            self.result_text_metadata.AppendText("Advertencia: No se encontraron metadatos o no se seleccionaron archivos válidos.")

    def display_directory_results(self, data):
        self.result_text_metadata.Clear()
        if data:
            for message in data:
                self.result_text_metadata.AppendText(f"{message}\n")
        else:
            self.result_text_metadata.AppendText("Advertencia: No se pudo eliminar los metadatos o la selección fue inválida.")

def analyze_metadata(self, filepath):
    try:
        if filepath.endswith('.pdf'):
            with open(filepath, 'rb') as file:
                pdf = PdfReader(file)
                if pdf.is_encrypted:
                    self.result_text_metadata.AppendText("Advertencia: El documento está firmado digitalmente. No se puede analizar.")
                info = pdf.metadata
                return info
        elif filepath.endswith('.docx'):
            doc = Document(filepath)
            metadata = {}
            props = doc.core_properties
            return {
                "Identificador": props.identifier or "N/A",
                "Título": props.title or "N/A",
                "Tema": props.subject or "N/A",
                "Autor": props.author or "N/A",
                "Última modificación": props.last_modified_by or "N/A",
                "Fecha de creación": props.created or "N/A",
                "Última modificación": props.modified or "N/A",
                "Categoría": props.category or "N/A",
                "Idioma": props.language or "N/A",
                "Estado del contenido": props.content_status or "N/A",
                "Palabras clave": props.keywords or "N/A",
                "Revisión": props.revision or "N/A",
                "Última impresión": props.last_printed or "N/A",
                "Comentarios": props.comments or "N/A",
                "Versión": props.version or "N/A"
            }
        elif filepath.endswith('.xlsx'):
            workbook = load_workbook(filepath, read_only=True)
            props = workbook.properties
            return {
                "Identificador": props.identifier or "N/A",
                "Título": props.title or "N/A",
                "Tema": props.subject or "N/A",
                "Descripción": props.description or "N/A",
                "Autor": props.creator or "N/A",
                "Última modificación": props.lastModifiedBy or "N/A",
                "Fecha de creación": props.created or "N/A",
                "Última modificación": props.modified or "N/A",
                "Categoría": props.category or "N/A",
                "Idioma": props.language or "N/A",
                "Estado del contenido": props.contentStatus or "N/A",
                "Palabras clave": props.keywords or "N/A",
                "Revisión": props.revision or "N/A",
                "Última impresión": props.lastPrinted or "N/A",
                "Versión": props.version or "N/A"
            }
        elif filepath.endswith('.pptx'):
            presentation = Presentation(filepath)
            props = presentation.core_properties
            return {
                "Identificador": props.identifier or "N/A",
                "Título": props.title or "N/A",
                "Tema": props.subject or "N/A",
                "Autor": props.author or "N/A",
                "Última modificación": props.last_modified_by or "N/A",
                "Fecha de creación": props.created or "N/A",
                "Última modificación": props.modified or "N/A",
                "Categoría": props.category or "N/A",
                "Idioma": props.language or "N/A",
                "Estado del contenido": props.content_status or "N/A",
                "Tipo de contenido": props.content_type or "N/A",
                "Palabras clave": props.keywords or "N/A",
                "Revisión": props.revision or "N/A",
                "Última impresión": props.last_printed or "N/A",
                "Comentarios": props.comments or "N/A",
                "Versión": props.version or "N/A"
            }
        elif filepath.endswith(('.jpg', '.jpeg', '.png')):
            image = Image.open(filepath)
            metadata = image.info
            return metadata
        elif filepath.endswith('.zip'):
            estadisticas = os.stat(filepath)
            print(estadisticas)
            metadata = {
                "Ruta": os.path.abspath(filepath) or "N/A",
                "Tamaño": estadisticas.st_size or "N/A",
                "Fecha de creación": datetime.fromtimestamp(estadisticas.st_birthtime).strftime("%Y-%m-%d %H:%M:%S") or "N/A",
                "Última modificación": datetime.fromtimestamp(estadisticas.st_mtime).strftime("%Y-%m-%d %H:%M:%S") or "N/A",
                "Último acceso": datetime.fromtimestamp(estadisticas.st_atime).strftime("%Y-%m-%d %H:%M:%S") or "N/A",
                "Modo permisos": estadisticas.st_mode or "N/A",  
                "Número inodo": estadisticas.st_ino or "N/A",  
                "Dispositivo": estadisticas.st_dev or "N/A",  
                "Número enlaces": estadisticas.st_nlink or "N/A",  
                "Propietario UID": estadisticas.st_uid or "N/A",
                "Grupo GID": estadisticas.st_gid or "N/A"
            }
            return metadata
        elif filepath.endswith(('.mp3', '.flac', '.wav', '.ogg')):
            audio = MutagenFile(filepath)
            return audio.tags if audio else "No tags found"
        elif filepath.endswith(('.mp4', '.mkv', '.avi', '.mov')):
            parser = createParser(filepath)
            if not parser:
                return "Unable to parse video file"
            metadata = extractMetadata(parser)
            return metadata.exportDictionary() if metadata else "No metadata found"
        
    except Exception as e:
        self.result_text_metadata.AppendText(f"ERROR: Error inesperado al analizar los metadatos: {e}")


def remove_metadata_pdf(self, filepath):
    try:
        reader = PdfReader(filepath)
        if reader.is_encrypted:
            self.result_text_metadata.AppendText("Advertencia: El documento PDF está firmado digitalmente. No se pueden eliminar los metadatos.")
        writer = PdfWriter()
        for page in range(len(reader.pages)):
            writer.add_page(reader.pages[page])
        writer.add_metadata({})
        with open(filepath, "wb") as f:
            writer.write(f)    
    except Exception as e:
        self.result_text_metadata.AppendText(f"ERROR: Error inesperado al eliminar metadatos del PDF: {e}")


def remove_metadata_audio(filepath):
    audio = MutagenFile(filepath, easy=True)
    if not audio:
        return f"No metadata found in {filepath}."
    
    audio.delete()
    audio.save()


def remove_metadata_video(filepath):

    parser = createParser(filepath)
    if not parser:
        return f"Unable to parse video file {filepath}."
    
    editor = createEditor(parser)
    if not editor:
        return f"Unable to create editor for {filepath}."
    
    for field in list(editor.iterFields()):
        editor.removeField(field)
    
    output_filepath = filepath.replace(".mp4", "_cleaned.mp4")  # Example for .mp4
    with open(output_filepath, "wb") as output_file:
        stream = FileOutputStream(output_file)
        editor.writeInto(stream)

def remove_metadata_office(self, filepath):
    temp_dir = "temp_file"
    file_extension = filepath.split('.')[-1]
    
    with zipfile.ZipFile(filepath, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)
    
    metadata_files = {
        'docx': ['docProps/core.xml', 'docProps/app.xml'],
        'xlsx': ['docProps/core.xml', 'docProps/app.xml'],
        'pptx': ['docProps/core.xml', 'docProps/app.xml']
    }
    
    if file_extension not in metadata_files:
        self.result_text_metadata.AppendText("Formato de archivo no soportado para eliminación de metadatos")
    
    for meta_file in metadata_files[file_extension]:
        meta_path = os.path.join(temp_dir, meta_file)
        if os.path.exists(meta_path):
            tree = ET.parse(meta_path)
            root = tree.getroot()
            for elem in root.iter():
                elem.clear()
            tree.write(meta_path)

    with zipfile.ZipFile(filepath, 'w') as zip_ref:
        for folder_name, subfolders, filenames in os.walk(temp_dir):
            for filename in filenames:
                file_path = os.path.join(folder_name, filename)
                zip_ref.write(file_path, os.path.relpath(file_path, temp_dir))

    shutil.rmtree(temp_dir)

def remove_metadata_image(filepath):
    image = Image.open(filepath)
    info = image.info
    if info:
        image.info.clear()
    image.save(filepath)

def remove_metadata_file(self, filepath):
    try:
        extension = os.path.splitext(filepath)[1]
        if extension in ['.jpg', '.jpeg', '.png']:
            remove_metadata_image(filepath)
            return f"Archivo: {os.path.basename(filepath)} - Los metadatos se eliminaron correctamente."
        elif extension == '.pdf':
            remove_metadata_pdf(self, filepath)
            return f"Archivo: {os.path.basename(filepath)} - Los metadatos se eliminaron correctamente."
        elif extension in ['.docx', '.xlsx', '.pptx']:
            remove_metadata_office(self, filepath)
            return f"Archivo: {os.path.basename(filepath)} - Los metadatos se eliminaron correctamente."
        elif extension in ['.mp3', '.flac', '.wav', '.ogg']:
            remove_metadata_audio(filepath)
            return f"File: metadata from {os.path.basename(filepath)} has been removed correctly."
        elif extension in ['.mp4', '.mkv', '.avi', '.mov']:
            remove_metadata_video(filepath)
            return f"File: metadata from {os.path.basename(filepath)} has been removed correctly."
        else:
            self.result_text_metadata.AppendText(f"El programa no soporta la eliminación de metadatos para la extensión: {extension}")
    except Exception as e:
        self.result_text_metadata.AppendText(f"ERROR: Error inesperado al eliminar los metadatos: {e}")


def remove_metadata_directory(self, file_list):
    try:
        info_list = []
        for file in file_list:
            info_list.append(remove_metadata_file(self, file))
        return info_list
    except Exception as e:
        self.result_text_metadata.AppendText(f"ERROR: Error removing metadata from directory: {e}")

def analyze_metadata_directory(self, file_list):
    try:
        info_list = []
        for file_read in file_list:
            info = analyze_metadata(self, file_read)
            info_list.append({
                "filename": file_read,
                "metadata": info
            })
        return info_list
    except Exception as e:
        self.result_text_metadata.AppendText(f"ERROR: Error analyzing files in directory: {e}")

def clear_results_metadata(result_text):
    result_text.Clear()
    wx.MessageBox("Text area has been cleared.", "Results cleared", wx.OK | wx.ICON_INFORMATION)


if __name__ == "__main__":
    app = MainApp()
    app.MainLoop()
