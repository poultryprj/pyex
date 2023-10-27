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



# Import the necessary module


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
                    # Add default columns to the new sheet
                    new_sheet.append(default_columns)

                    # Set column widths for default columns
                    for column_letter, column_name in zip('ABCDEFGHIJKLM', default_columns):
                        column = new_sheet.column_dimensions[column_letter]
                        # Adjust the width as needed
                        column.width = len(column_name) + 2

                    # Set the alignment for the header row (centered)
                    header_row = new_sheet[1]
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
                    # Add default columns to the new sheet
                    new_sheet.append(default_columns)

                    # Set column widths for default columns
                    for column_letter, column_name in zip('ABCDEFGHIJKLM', default_columns):
                        column = new_sheet.column_dimensions[column_letter]
                        # Adjust the width as needed
                        column.width = len(column_name) + 2

                    # Set the alignment for the header row (centered)
                    header_row = new_sheet[1]
                    for cell in header_row:
                        cell.alignment = Alignment(horizontal='center')

                sheet = new_sheet
            else:
                sheet = workbook[sheet_name]

            # Check if column names already exist in the sheet
            column_names = [cell.value for cell in sheet[1]]

            # Get the current date and time in UTC+05:30 (IST)
            # Define the IST timezone
            ist = pytz.timezone('Asia/Kolkata')

            # Get the current date and time in the IST timezone
            current_datetime = datetime.now(ist)

            # Format current_date as "dd/mm/yyyy" and current_time as "HH:mm:ss"
            current_date = current_datetime.strftime(
                '%d/%m/%Y')  # Format date as dd/mm/yyyy
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
                    if product_id == 1:
                        amount= obj.get('weight', 0.0) * obj.get('daily_rate', 0.0)
                    else:
                        amount= obj.get('quantity', 0.0) * obj.get('daily_rate', 0.0)


                    try:
                        excel_data = ExcelData(
                            date=current_datetime.now(),
                            time=current_datetime.ctime(),
                            shop_code=obj.get('shop_code', 0),
                            product_id=product_id,
                            weight=weight,  # Use the value set above
                            quantity=obj.get('quantity', 0.0),
                            daily_rate=obj.get('daily_rate', 0.0),
                            rate=obj.get('rate', 0.0),                            
                            amount = amount
                        )
                        excel_data.save()
                    except Exception as e:
                        print(e)

                    # Handle the weight when appending to the sheet
                    if weight == "":
                        print(excel_data.amount,"amount")
                        row = [current_date, current_time, obj.get('shop_code', 0), product_id, "", obj.get(
                            'quantity', 0.0), obj.get('daily_rate', 0.0), obj.get('rate', 0.0), excel_data.amount]
                    else:
                        row = [current_date, current_time, obj.get('shop_code', 0), product_id, weight, obj.get(
                            'quantity', 0.0), obj.get('daily_rate', 0.0), obj.get('rate', 0.0), excel_data.amount]

                    # Append data to the sheet
                    sheet.append(row)

                    # Create an Alignment object to center align text
                    alignment = Alignment(horizontal='center')

                    # Apply the alignment to all cells in the last row of the sheet
                    for cell in sheet[sheet.max_row]:
                        cell.alignment = alignment

                # Save the updated Excel file
                workbook.save(file_path)
                return Response({'message': 'Data appended successfully'}, status=status.HTTP_200_OK)
            else:
                return Response({"error": "Column names in the data do not match the existing sheet"}, status=status.HTTP_400_BAD_REQUEST)
        except Exception as e:
            return Response({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)
        

    
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from django.http import JsonResponse
from rest_framework.decorators import api_view
from rest_framework import status
import json

@api_view(['POST'])
def create_daily_summary_sheet(request, sheet_name):
    if request.method == 'POST':
        try:
            data = json.loads(request.body.decode("utf-8"))

            # Provide the full path to your Excel file
            file_path = 'main.xlsx'

            # Define the default columns
            default_columns = [
                'Date',
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
                '',
                'opening_balance',
                'paid_amount',
                'closing_balance'
            ]

            # Load the Excel workbook using openpyxl
            workbook = load_workbook(filename=file_path)

            if sheet_name not in workbook.sheetnames:
                # Create a new sheet with the provided sheet_name
                new_sheet = workbook.create_sheet(title=sheet_name)

                # Add default columns to the new sheet
                new_sheet.append(default_columns)

                # Set column widths for default columns
                for column_letter, column_name in zip('ABCDEFGHIJKLMNOPQRST', default_columns):
                    column = new_sheet.column_dimensions[column_letter]
                    # Adjust the width as needed
                    column.width = max(len(column_name) + 2, 12)  # Minimum width of 12

                # Set the alignment for the header row (centered)
                header_row = new_sheet[1]
                for cell in header_row:
                    cell.alignment = Alignment(horizontal='center')

                sheet = new_sheet
            else:
                # Use the existing sheet with the provided sheet_name
                sheet = workbook[sheet_name]

            # Check if column names already exist in the sheet
            column_names = [cell.value for cell in sheet[1]]

            # Define the financial year start and end dates
            financial_year_start = datetime(2023, 4, 1)
            financial_year_end = datetime(2024, 3, 31)

            # Iterate over each date within the financial year
            current_date = financial_year_start
            while current_date <= financial_year_end:
                # Create a new row for each date
                row = [current_date.strftime('%d/%m/%Y'), 1, '', '', '', '', '', 2, '', '', '', '']
                sheet.append(row)

                # Move to the next date
                current_date += timedelta(days=1)

            # Save the updated Excel file
            workbook.save(file_path)
            return JsonResponse({'message': 'Data appended successfully'}, status=status.HTTP_200_OK)
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)


from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from django.http import JsonResponse
from rest_framework.decorators import api_view
import json
from rest_framework import status

@api_view(['POST'])
def insert_formulas_to_weight_column(request, sheet_name, file_path):
    if request.method == 'POST':
        try:
            data = json.loads(request.body.decode("utf-8"))

            # Load the Excel workbook using openpyxl
            workbook = load_workbook(filename=file_path)

            if sheet_name not in workbook.sheetnames:
                return JsonResponse({'error': 'Sheet does not exist'}, status=status.HTTP_400_BAD_REQUEST)

            sheet = workbook[sheet_name]

            # Check if the "weight" column exists, and if not, create it
            weight_column = None
            for col, column_name in enumerate(sheet.iter_rows(min_row=1, max_row=1, values_only=True)):
                col += 1  # Adjust for 1-based indexing
                for column in column_name:
                    if column == 'weight':
                        print(column)
                        weight_column = get_column_letter(col)

            if weight_column is None:
                return JsonResponse({'error': 'Weight column name does not exist'}, status=status.HTTP_400_BAD_REQUEST)

            # Define the financial year start and end dates
            financial_year_start = datetime(2023, 4, 1)
            financial_year_end = datetime(2024, 3, 31)

            # Iterate over each date within the financial year
            current_date = financial_year_start
            row_number = 2  # Start from the second row (1-based index)

            while current_date <= financial_year_end:
                # Append the formula to the "weight" column for each row
                date_cell = f"A{row_number}"
                formula = (
                    f'=IF(COUNTIFS(A:A,{date_cell},D:D,1)>0,'
                    f'AVERAGEIFS(E:E,A:A,{date_cell},D:D,1),"")'
                )
                sheet[f"{weight_column}{row_number}"] = formula

                # Move to the next date and row
                current_date += timedelta(days=1)
                row_number += 1

            # Save the updated Excel file
            workbook.save(file_path)
            return JsonResponse({'message': 'Formulas inserted successfully'})
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)
