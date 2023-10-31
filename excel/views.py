from django.shortcuts import render
import json
from openpyxl import load_workbook
from django.http import JsonResponse
from rest_framework.decorators import api_view
from openpyxl.utils.dataframe import dataframe_to_rows
from rest_framework import status
import pandas as pd
from openpyxl import load_workbook, Workbook
from rest_framework.response import Response
from .models import ExcelData  # Import your ExcelData model
from datetime import datetime  # Import datetime module
import pytz  # Import pytz module
from openpyxl.styles import Alignment
from datetime import datetime, timedelta
import openpyxl
from openpyxl.worksheet.views import SheetView, Selection
import os

@api_view(['POST'])
def excel_view(request, sheet_name):
    if request.method == 'POST':
        try:
            data = json.loads(request.body.decode("utf-8"))
            json_objects = data.get('json_objects')

            # Provide the full path to your Excel file
            file_path = 'main.xlsx'

            if not json_objects:
                return Response({"error": "JSON objects are required."}, status=status.HTTP_400_BAD_REQUEST)

            # Define the default columns outside of the if-else block
            default_columns = ['   date   ', '   time   ', 'shop_code',
                               'product_id', 'weight', 'quantity', 'daily_rate', 'rate', 'amount']

            # Load the Excel workbook using openpyxl
            workbook = load_workbook(filename=file_path)

            # Find the current sheet or create it if it doesn't exist
            if sheet_name not in workbook.sheetnames:
                if len(workbook.sheetnames) == 0:
                    # Create the first sheet
                    new_sheet_name = "Raw_data_01"
                    new_sheet = workbook.create_sheet(title=new_sheet_name)
                    new_sheet.append(['Raw Data Summary'])  # Add the title row
                    new_sheet.append(default_columns)  # Add the column names

                    # Set column widths for default columns
                    for column_letter, column_name in zip('ABCDEFGHIJKLM', default_columns):
                        column = new_sheet.column_dimensions[column_letter]
                        # Adjust the width as needed
                        column.width = len(column_name) + 2

                    # Set the alignment for the header row (centered)
                    header_row = new_sheet[2]
                    for cell in header_row:
                        cell.alignment = Alignment(horizontal='center')
                else:
                    # Find the next available sheet name with a sequential number
                    sheet_number = 1
                    while True:
                        new_sheet_name = f"Raw_data_{sheet_number:02d}"
                        if new_sheet_name not in workbook.sheetnames:
                            break
                        sheet_number += 1
                    new_sheet = workbook.create_sheet(title=new_sheet_name)
                    # Merge cells for the title
                    new_sheet.merge_cells('A1:I1')
                    title_cell = new_sheet.cell(row=1, column=1)
                    title_cell.value = 'RAW DATA'
                    title_cell.alignment = Alignment(horizontal='center')
                    new_sheet.append(default_columns)  # Add the column names

                    # Set column widths for default columns
                    for column_letter, column_name in zip('ABCDEFGHIJKLM', default_columns):
                        column = new_sheet.column_dimensions[column_letter]
                        # Adjust the width as needed
                        column.width = len(column_name) + 2

                    # Set the alignment for the header row (centered)
                    header_row = new_sheet[2]
                    for cell in header_row:
                        cell.alignment = Alignment(horizontal='center')

                sheet = new_sheet
            else:
                sheet = workbook[sheet_name]

            # Check if column names already exist in the sheet
            column_names = [cell.value for cell in sheet[2]]

            # Get the current date and time in UTC+05:30 (IST)
            # Define the IST timezone
            ist = pytz.timezone('Asia/Kolkata')

            # Get the current date and time in the IST timezone
            current_datetime = datetime.now(ist)

            # Format current_date as "dd/mm/yyyy" and current_time as "HH:mm:ss"
            current_date = current_datetime.strftime(
                '%d-%m-%Y')  # Format date as dd/mm/yyyy
            current_time = current_datetime.strftime(
                '%H:%M:%S')  # Format time as HH:mm:ss

            # Append data only if the column names match and product_id is 1 or 2
            if column_names == default_columns:
                # ...
                # Inside the loop that processes JSON objects
                for obj in json_objects:
                    product_id = obj.get('product_id', 0)
                    if product_id not in (1, 2):
                        return Response({"error": "Please enter a valid product_id (1 or 2)"}, status=status.HTTP_400_BAD_REQUEST)

                    if product_id == 2:  # product_id == 2  for bird and weight not needed
                        weight = obj.get('weight')
                        if weight:
                            return Response({"error": "please remove weight field..!!"}, status=status.HTTP_400_BAD_REQUEST)
                    else:
                        # product_id == 1  for eggs and weight needed
                        weight = obj.get('weight')
                        if not weight:
                            return Response({"error": "please enter weight..!!"}, status=status.HTTP_400_BAD_REQUEST)

                    # Handle the weight when appending to the sheet
                    if weight == "":
                        row = [current_date, current_time, obj.get('shop_code', 0), product_id, "", obj.get(
                            'quantity', 0.0), obj.get('daily_rate', 0.0), obj.get('rate', 0.0), ""]
                    else:
                        row = [current_date, current_time, obj.get('shop_code', 0), product_id, weight, obj.get(
                            'quantity', 0.0), obj.get('daily_rate', 0.0), obj.get('rate', 0.0), ""]

                    # Append data to the sheet
                    sheet.append(row)
                    sheet.freeze_panes = "A3"

                    # Create an Alignment object to center align text
                    alignment = Alignment(horizontal='center')

                    # Apply the alignment to all cells in the last row of the sheet
                    for cell in sheet[sheet.max_row]:
                        cell.alignment = alignment

                # Save the updated Excel file after appending data
                workbook.save(file_path)

                # Apply the formula after saving the sheet
                for i in range(3, sheet.max_row + 1):
                    formula = f'=IF(D{i}=1, E{i}*G{i}, IF(D{i}=2, F{i}*G{i}, ""))'
                    sheet.cell(row=i, column=default_columns.index(
                        "amount") + 1).value = formula

                # Save the Excel file again after applying the formula
                workbook.save(file_path)

                return Response({'message': 'Data appended successfully'}, status=status.HTTP_200_OK)
            else:
                return Response({"error": "Column names in the data do not match the existing sheet"}, status=status.HTTP_400_BAD_REQUEST)
        except Exception as e:
            return Response({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)



#################################@api_view(['POST'])
@api_view(['POST'])
def create_daily_summary_sheet(request, sheet_name):
    if request.method == 'POST':
        try:
            data = json.loads(request.body.decode("utf-8"))

            # Provide the full path to your Excel file
            file_path = 'main.xlsx'

            # Check if the file exists
            if not os.path.isfile(file_path):
                return JsonResponse({'error': f'File not found at path: {file_path}'}, status=status.HTTP_400_BAD_REQUEST)

            # Load the Excel workbook using openpyxl
            workbook = load_workbook(filename=file_path)

            if sheet_name in workbook.sheetnames:
                return JsonResponse({'error': f'Sheet "{sheet_name}" already exists'}, status=status.HTTP_400_BAD_REQUEST)

            # Create a new sheet with the provided sheet_name
            new_sheet = workbook.create_sheet(title=sheet_name)

            # Define the default columns and add them to the A2 row
            default_columns = [
                'Date',
                'opening_balance',
                'paid_amount',
                'closing_balance',
                '',
                'product_1',
                'weight',
                'quantity',
                'rate',
                'amount',
                '',
                'product_2',
                'quantity',
                'rate',
                'amount',
            ]

            # Merge cells for the title
            new_sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
            title_cell = new_sheet.cell(row=1, column=1)
            title_cell.value = 'DAILY ACCOUNT SUMMARY'
            title_cell.alignment = Alignment(horizontal='center')

            new_sheet.merge_cells(start_row=1, start_column=6, end_row=1, end_column=10)
            title_cell = new_sheet.cell(row=1, column=6)
            title_cell.value = 'BIRDS'
            title_cell.alignment = Alignment(horizontal='center')

            new_sheet.merge_cells(start_row=1, start_column=12, end_row=1, end_column=15)
            title_cell = new_sheet.cell(row=1, column=12)
            title_cell.value = 'EGGS'
            title_cell.alignment = Alignment(horizontal='center')

            for col_num, header in enumerate(default_columns, start=1):
                new_sheet.cell(row=2, column=col_num, value=header)

            # Set the alignment for the header row (centered)
            header_row = new_sheet[2]
            for cell in header_row:
                cell.alignment = Alignment(horizontal='center')

            # Set column widths for default columns
            for column_letter, column_name in zip('ABCDEFGHIJKLMNOPQRST', default_columns):
                column = new_sheet.column_dimensions[column_letter]
                # Adjust the width as needed
                # Minimum width of 12
                column.width = max(len(column_name) + 2, 12)

            # Define the financial year start and end dates
            financial_year_start = datetime(2023, 4, 1)
            financial_year_end = datetime(2024, 3, 31)

            # Iterate over each date within the financial year
            current_date = financial_year_start
            while current_date <= financial_year_end:
                # Create a new row for each date
                row = [current_date.strftime('%d-%m-%Y'), '', '', '', '', 1, '', '', '', '', '', 2, '', '', '']
                new_sheet.append(row)

                # Move to the next date
                current_date += timedelta(days=1)

            # Add the formulas to the "G" column (weight column) from $A3 to $A368
            for row in range(3, 369):
                # For product_id 1
                avg_weight = f'=IF(COUNTIFS(Raw_data_01!A:A,$A{row},Raw_data_01!D:D,1)>0,AVERAGEIFS(Raw_data_01!E:E,Raw_data_01!A:A,$A{row},Raw_data_01!D:D,1),"")'
                new_sheet[f'G{row}'] = avg_weight

                sum_of_quantity = f'=IF(COUNTIFS(Raw_data_01!A:A,$A{row},Raw_data_01!D:D,1)>0,SUMIFS(Raw_data_01!F:F,Raw_data_01!A:A,$A{row},Raw_data_01!D:D,1),"")'
                new_sheet[f'H{row}'] = sum_of_quantity

                avg_rate = f'=IF(COUNTIFS(Raw_data_01!A:A,$A{row},Raw_data_01!D:D,1)>0,AVERAGEIFS(Raw_data_01!H:H,Raw_data_01!A:A,$A{row},Raw_data_01!D:D,1),"")'
                new_sheet[f'I{row}'] = avg_rate

                sum_of_amount = f'=IF(COUNTIFS(Raw_data_01!A:A,$A{row},Raw_data_01!D:D,1)>0,SUMIFS(Raw_data_01!I:I,Raw_data_01!A:A,$A{row},Raw_data_01!D:D,1),"")'
                new_sheet[f'J{row}'] = sum_of_amount

                # For product_id 2
                sum_of_quantity = f'=IF(COUNTIFS(Raw_data_01!A:A,$A{row},Raw_data_01!D:D,2)>0,SUMIFS(Raw_data_01!F:F,Raw_data_01!A:A,$A{row},Raw_data_01!D:D,2),"")'
                new_sheet[f'M{row}'] = sum_of_quantity

                avg_rate = f'=IF(COUNTIFS(Raw_data_01!A:A,$A{row},Raw_data_01!D:D,2)>0,AVERAGEIFS(Raw_data_01!H:H,Raw_data_01!A:A,$A{row},Raw_data_01!D:D,2),"")'
                new_sheet[f'N{row}'] = avg_rate

                sum_of_amount = f'=IF(COUNTIFS(Raw_data_01!A:A,$A{row},Raw_data_01!D:D,2)>0,SUMIFS(Raw_data_01!I:I,Raw_data_01!A:A,$A{row},Raw_data_01!D:D,2),"")'
                new_sheet[f'O{row}'] = sum_of_amount

                closing_balance = f'=SUM(J{row},O{row},B{row}) - C{row}'
                new_sheet[f'D{row}'] = closing_balance

            # Add the formula to the "B" column (opening_balance column) from B4 to B368
            for row in range(4, 369):
                formula = f'=IF(D{row - 1}<>0, D{row - 1}, IFERROR(INDEX(D3:D${row - 1}, MATCH(1, D3:D${row - 1}<>0, 0)), LOOKUP(2, 1/(D3:D${row - 1}<>0), D3:D${row - 1})))'
                new_sheet[f'B{row}'] = formula

            # Format the columns
            columns_to_format = ['G', 'H', 'I', 'J', 'M', 'N', 'O', 'B', 'C', 'D']
            for col_letter in columns_to_format:
                # Format the columns to display two decimal places
                for row in new_sheet.iter_rows(min_row=3, max_row=369, min_col=ord(col_letter) - 64, max_col=ord(col_letter) - 64):
                    for cell in row:
                        cell.number_format = '0.00'

            # Freeze the top row (column names) when scrolling
            new_sheet.freeze_panes = "A3"

            # Save the updated Excel file again
            workbook.save(file_path)

            return JsonResponse({'message': f'Successfully Created {sheet_name} with Data & Formulas'}, status=status.HTTP_200_OK)
        except FileNotFoundError as e:
            return JsonResponse({'error': 'File not found'}, status=status.HTTP_400_BAD_REQUEST)
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)
    else:
        return JsonResponse({'error': 'Invalid request method'}, status=status.HTTP_400_BAD_REQUEST)
