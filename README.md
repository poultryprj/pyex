```markdown
# Project Documentation

This documentation provides information about a Django project that offers RESTful API endpoints for managing data in an Excel file.

## Setup and Installation

1. Clone the repository:

   ```shell
   git clone https://github.com/your/repository.git
   cd your-repository
   ```

2. Create a virtual environment and install the dependencies:

   ```shell
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   pip install -r requirements.txt
   ```

3. Configure your database settings in `settings.py`. For this example, we're using SQLite as the default database.

4. Apply database migrations:

   ```shell
   python manage.py migrate
   ```

5. Start the Django development server:

   ```shell
   python manage.py runserver
   ```

## API Endpoints

### Create Excel Data

This endpoint allows you to add data to an Excel file. The Excel file is expected to have a sheet named 'main.xlsx' to store the data.

- **URL:** `/create_excel/<str:sheet_name>/`

- **Method:** `POST`

- **Request Body:**

  ```json
  {
      "json_objects": [
          {
              "shop_code": "A001",
              "product_type": "ACCOUNT",
              "product_id": 1,
              "weight": 10.5,
              "quantity": 100,
              "daily_rate": 150.0,
              "rate": 15.0
          },
          // Add more JSON objects as needed
      ]
  }
  ```

- **Response:**

  - `200 OK` on successful data insertion.

  - `400 Bad Request` on invalid input data.

### Create Daily Summary Sheet

This endpoint generates a daily summary sheet in the 'main.xlsx' Excel file. The summary sheet contains formulas to calculate average weight, total quantity, average rate, and total amount for different product types.

- **URL:** `/create_daily_summary_sheet/<str:sheet_name>/`

- **Method:** `POST`

- **Request Body:**

  ```json
  {}  // Empty request body
  ```

- **Response:**

  - `200 OK` on successful summary sheet creation.

  - `400 Bad Request` if the sheet already exists or an error occurs.

## Usage Example with Postman

1. Start your Django server:

   ```shell
   python manage.py runserver
   ```

2. Use Postman to test the API endpoints.

   - **Create Excel Data:**

     - URL: `http://localhost:8000/create_excel/sheet_name/`
     - Method: `POST`
     - Request Body: Provide a JSON object as shown in the "Request Body" section above.
     - Check the response to ensure that the data is successfully inserted.

   - **Create Daily Summary Sheet:**

     - URL: `http://localhost:8000/create_daily_summary_sheet/sheet_name/`
     - Method: `POST`
     - Request Body: An empty JSON object.
     - Check the response to confirm that the summary sheet is created with formulas.

## Excel File Structure

- The Excel file should have a sheet named 'main.xlsx' for data storage and summary sheet creation.

- The 'main.xlsx' sheet should have the following columns: date, time, shop_code, product_type, product_id, weight, quantity, daily_rate, rate, and amount.

- The 'main.xlsx' sheet should follow the defined structure for different product types as described in the code.

## Notes

- Please adjust the database settings, such as database type and connection details, in the project's `settings.py`.

- For production use, consider deploying the project on a production-ready web server and database.

- This documentation assumes that you have a basic understanding of Python and Django.

Feel free to adapt this README.md to your project's specific requirements and add any additional information or instructions that your users may need. Make sure to keep it updated as your project evolves.