import os
import pikepdf


def flatten_with_pikepdf(input_path, output_path):

    if not os.path.exists(input_path):
        return

    try:        
        with pikepdf.Pdf.open(input_path) as pdf:
            pdf.flatten_annotations(mode='all')
            if '/AcroForm' in pdf.Root:
                 del pdf.Root['/AcroForm']
            pdf.save(output_path)
    except Exception as e:
        print("Unexpected error while flattening PDF")
if __name__ == "__main__":
    flatten_with_pikepdf(input_file, output_file)