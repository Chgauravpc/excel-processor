import os
import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter
import tempfile
import requests
from vercel_blob import put, get

def handler(request):
    try:
        data = request.json
        if 'filename' in data:
            # Generate presigned URL for upload
            filename = data['filename']
            if not filename.endswith('.xlsx'):
                return {
                    'statusCode': 400,
                    'body': {'error': 'File must be .xlsx'}
                }
            key = f'uploads/{filename}'
            upload_url = put(key, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', generate_url=True)
            return {
                'statusCode': 200,
                'body': {'uploadUrl': upload_url, 'key': key}
            }
        elif 'key' in data:
            # Process the uploaded file
            key = data['key']
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_input:
                # Download file from Vercel Blob
                response = requests.get(get(key))
                response.raise_for_status()
                temp_input.write(response.content)
                temp_input_path = temp_input.name

            # Process the file (your script)
            workbook = openpyxl.load_workbook(temp_input_path)
            worksheet = workbook.active
            max_row = worksheet.max_row

            # Set row heights
            for row in range(2, 8):
                worksheet.row_dimensions[row].height = 15
            for row in range(8, 13):
                worksheet.row_dimensions[row].height = 0
                worksheet.row_dimensions[row].hidden = True
            for row in range(14, max_row - 1):
                worksheet.row_dimensions[row].height = 27

            # Set column widths
            column_widths = {
                'B': 3, 'E': 9.8, 'F': 13, 'G': 26.3, 'I': 6, 'J': 0,
                'L': 3.2, 'M': 2.7, 'N': 3, 'Q': 2.7, 'U': 7.3, 'V': 5, 'W': 5.5
            }
            for col, width in column_widths.items():
                worksheet.column_dimensions[col].width = width
                if width == 0:
                    worksheet.column_dimensions[col].hidden = True
            columns_to_hide = ['C', 'P', 'R', 'S', 'T', 'X', 'Y']
            for col in columns_to_hide:
                worksheet.column_dimensions[col].width = 0
                worksheet.column_dimensions[col].hidden = True

            # Define styles
            border_style = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
            font_style = Font(size=10)
            font_style_max_row = Font(size=16, bold=True)

            # Format B13:W(max_row - 2)
            for row in range(13, max_row - 1):
                for col_idx in range(2, 24):
                    col_letter = get_column_letter(col_idx)
                    cell = worksheet[f'{col_letter}{row}']
                    cell.border = border_style
                    cell.font = font_style
                    if col_idx >= 5 and row >= 14:
                        cell.alignment = Alignment(horizontal='left', vertical='center', wrapText=True)
                    else:
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrapText=True)
                    if col_idx == 2 and row >= 14:
                        cell.value = row - 13

            # Merge K(max_row - 1):W(max_row - 1)
            if max_row >= 2:
                merge_row = max_row - 1
                worksheet.merge_cells(f'K{merge_row}:W{merge_row}')
                merged_cell = worksheet[f'K{merge_row}']
                merged_cell.font = font_style
                merged_cell.alignment = Alignment(horizontal='left', vertical='center', wrapText=True)
                merged_cell.border = border_style
                if merge_row < 14:
                    worksheet.row_dimensions[merge_row].height = 27

            # Apply font size 16 and bold to row max_row
            if max_row >= 1:
                for col_idx in range(1, worksheet.max_column + 1):
                    col_letter = get_column_letter(col_idx)
                    cell = worksheet[f'{col_letter}{max_row}']
                    cell.font = font_style_max_row
                    if max_row < 14:
                        worksheet.row_dimensions[max_row].height = 27

            # Save processed file to temporary location
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_output:
                workbook.save(temp_output.name)
                temp_output_path = temp_output.name

            # Upload processed file to Vercel Blob
            output_key = f'processed/processed_{os.path.basename(key)}'
            with open(temp_output_path, 'rb') as f:
                download_url = put(output_key, f.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', generate_url=True)

            # Clean up temporary files
            os.unlink(temp_input_path)
            os.unlink(temp_output_path)

            return {
                'statusCode': 200,
                'body': {'downloadUrl': download_url}
            }
        else:
            return {
                'statusCode': 400,
                'body': {'error': 'Invalid request: provide filename or key'}
            }
    except Exception as e:
        return {
            'statusCode': 500,
            'body': {'error': str(e)}
        }