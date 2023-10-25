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

# Import the necessary module


@api_view(['POST'])
def excel_view(request, sheet_name):
    if request.method == 'POST':
        sheet_name_trimed = sheet_name[:8]
        if "Raw_data_" == sheet_name_trimed:
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
            
        ### if sheetname Summary_data_01
        sheet_name_trimed = sheet_name[:13]
        if "Summary_data_" == sheet_name_trimed:
            try:
                data = json.loads(request.body.decode("utf-8"))
                json_objects = data.get('json_objects')

                # Provide the full path to your Excel file
                file_path = 'main.xlsx'

                if not json_objects:
                    return Response({"error": "JSON objects are required."}, status=status.HTTP_400_BAD_REQUEST)

                # Define the default columns outside of the if-else block
                default_columns = ['shop_code', '   date   ','opening_balance', 
                                   'product_id', 'weight', 'quantity', 'daily_rate', 'rate', 'amount', 'paid_amount', 'closing_balance', 'average_rate', 'sum_of_weight', 'sum_of_quantity']

                # Load the Excel workbook using openpyxl
                workbook = load_workbook(filename=file_path)

                # Find the current sheet or create it if it doesn't exist
                if sheet_name not in workbook.sheetnames:
                    if len(workbook.sheetnames) == 0:
                        # Create the first sheet
                        new_sheet_name = "Summary_data_01"
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
                            new_sheet_name = f"Summary_data_{sheet_number:02d}"
                            if new_sheet_name not in workbook.sheetnames:
                                break
                            sheet_number += 1
                        new_sheet = workbook.create_sheet(title=new_sheet_name)
                        # Add default columns to the new sheet
                        new_sheet.append(default_columns)

                        # Set column widths for default columns
                        for column_letter, column_name in zip('ABCDEFGHIJKLMNOPQRST', default_columns):
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
                                shop_code=obj.get('shop_code', 0),
                                date=current_datetime.now(),
                                opening_balance = obj.get('opening_balance', 0.0),
                                product_id=product_id,
                                weight=weight,  # Use the value set above
                                quantity=obj.get('quantity', 0.0),
                                daily_rate=obj.get('daily_rate', 0.0),
                                rate=obj.get('rate', 0.0),                            
                                amount = amount,
                                paid_amount = obj.get('paid_amount', 0.0),
                                closing_balance = obj.get('closing_balance', 0.0)
                            )
                            excel_data.save()
                        except:
                            pass

                        # Handle the weight when appending to the sheet
                        if weight == "":
                            print(excel_data.amount,"amount")
                            row = [ obj.get('shop_code', 0), current_date, excel_data.opening_balance, product_id, "", obj.get(
                                'quantity', 0.0), obj.get('daily_rate', 0.0), obj.get('rate', 0.0), excel_data.amount, excel_data.paid_amount, excel_data.closing_balance]
                        else:
                            row = [ obj.get('shop_code', 0), current_date, excel_data.opening_balance, product_id, weight, obj.get(
                                'quantity', 0.0), obj.get('daily_rate', 0.0), obj.get('rate', 0.0), excel_data.amount, excel_data.paid_amount, excel_data.closing_balance]

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
        else:
            return Response({"error": "please enter valid excel sheet name..!!"}, status=status.HTTP_400_BAD_REQUEST)

@api_view(['GET'])
def read_excel(request, sheet_name):
    try:
        # Provide the full path to your Excel file
        file_path = r'main.xlsx'

        # Load the Excel workbook using openpyxl
        workbook = load_workbook(filename=file_path)

        if sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]

            excel_data = []

            # Extract data from the Excel sheet
            for row in sheet.iter_rows(min_row=2, values_only=True):
                sr_no, name, amount, pending = row
                excel_data.append({
                    "sr_no": sr_no,
                    "name": name,
                    "amount": amount,
                    "pending": pending
                })

            return JsonResponse({'data': excel_data}, status=status.HTTP_200_OK)
        else:
            return JsonResponse({'error': 'Sheet not found'}, status=status.HTTP_404_NOT_FOUND)
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)


@api_view(['PUT'])
def update_excel(request, sheet_name):
    if request.method == 'PUT':
        try:
            data = json.loads(request.body.decode("utf-8"))
            json_objects = data.get('json_objects')

            # Provide the full path to your Excel file
            file_path = r'main.xlsx'

            if not json_objects:
                return JsonResponse({"error": "JSON objects are required."}, status=status.HTTP_400_BAD_REQUEST)

            # Load the Excel workbook using openpyxl
            workbook = load_workbook(filename=file_path)

            if sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]

                # Iterate through JSON objects and update the Excel file
                for obj in json_objects:
                    sr_no = obj.get('sr_no')
                    name = obj.get('name')
                    amount = obj.get('amount')
                    pending = obj.get('pending')

                    if sr_no is None:
                        return JsonResponse({'error': 'sr_no is mandatory for updates'}, status=status.HTTP_400_BAD_REQUEST)

                    found = False

                    # Find and update the row with the matching sr_no
                    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                        if row[0].value == sr_no:
                            row[1].value = name
                            row[2].value = amount
                            row[3].value = pending
                            found = True
                            break

                    if not found:
                        return JsonResponse({'error': f"Row with sr_no {sr_no} not found"}, status=status.HTTP_404_NOT_FOUND)

                # Save the updated Excel file
                workbook.save(file_path)

                return JsonResponse({'message': 'Data updated successfully'}, status=status.HTTP_200_OK)
            else:
                return JsonResponse({'error': 'Sheet not found'}, status=status.HTTP_404_NOT_FOUND)

        except Exception as e:
            return JsonResponse({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)


@api_view(['DELETE'])
def delete_excel(request, sheet_name):
    if request.method == 'DELETE':
        try:
            # Provide the full path to your Excel file
            file_path = r'main.xlsx'

            # Load the Excel workbook using openpyxl
            workbook = load_workbook(filename=file_path)

            if sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]

                # Get the sr_no to be deleted from the request data
                data = json.loads(request.body.decode("utf-8"))
                sr_no_to_delete = data.get('sr_no')

                if sr_no_to_delete is None:
                    return JsonResponse({'error': 'sr_no is required for row deletion'}, status=status.HTTP_400_BAD_REQUEST)

                found = False

                # Find and delete the row with the matching sr_no
                for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                    if row[0].value == sr_no_to_delete:
                        sheet.delete_rows(row[0].row)
                        found = True
                        break

                if not found:
                    return JsonResponse({'error': f"Row with sr_no {sr_no_to_delete} not found"}, status=status.HTTP_404_NOT_FOUND)

                # Save the updated Excel file
                workbook.save(file_path)

                return JsonResponse({'message': 'Row deleted successfully'}, status=status.HTTP_200_OK)
            else:
                return JsonResponse({'error': 'Sheet not found'}, status=status.HTTP_404_NOT_FOUND)
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)
