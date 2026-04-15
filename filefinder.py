from pathlib import Path

import shutil

source=Path(r"C:/Users/Admins/Downloads/")

dest=Path(r"C:/Users/Admins/Documents/")
dest.mkdir(parents=True, exist_ok=True)
print("Hiiiiiiiiiiiiiiiii")

filesignature={
    b'%PDF': 'PDF',
    b'\x89PNG\r\n\x1a\n': 'PNG',
    b'\xff\xd8\xff': 'JPG',
    b'GIF8': 'GIF',
    b'PK\x03\x04': 'ZIP/DOCX/XLSX/PPTX',
    b'Rar!': 'RAR',
    b'MZ': 'EXE',
    b'ID3': 'MP3',
    b'ftyp': 'MP4',
}

def filename(file):
        with open(file, 'rb') as f:
            byte=f.read(8) 
        
        for sig, name in filesignature.items():
            print(byte, sig )
            if byte.startswith(sig):  
                return name
        return "UNKNOWN"
        
                    
def store_file_signature(source):
     for file in source.glob('*'):
          if file.is_file():
            type=filename(file)    
           

            destfolder=dest/type
            destfolder.mkdir(parents=True, exist_ok=True)  
            shutil.move(str(file), destfolder/file.name)

store_file_signature(source)




