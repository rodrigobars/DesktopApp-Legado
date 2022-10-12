import win32com.client as win32
import os

# Limpando o gen.py antes de ocorrer erro
os.system("powershell Remove-Item -path $env:LOCALAPPDATA\Temp\gen_py -recurse")
os.system("cls")

ata_path = input("\nInsira o caminho da Ata: ")

os.system("cls")

# Abrindo o programa Word e setando a visibilidade como verdadeira
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = True

# Abrindo o caminho do documento Word
wordDoc = word.Documents.Open(rf'{ata_path}')

actual_table = wordDoc.Tables(1)
actual_table.Cell(1, 0).Range.Select()
word.Selection.Paragraphs.Alignment = 0
#print(dir(word.Selection.Paragraphs(1).Alignment()))

#C:\Users\Rodrigo.DESKTOP-ST2DND8\Desktop\Work\Atas\2022\1. PENDENTE\2. Necessito elaborar a Ata\teste\teste.docx