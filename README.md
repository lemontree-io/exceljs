# excelJS - Create Excel files from HTML with JavaScript
excelJS is a JavaScript framework for generating Excel *.xlsx files straight from your html tables. 


Getting Started
------
1. Download file *exceljs.compiled.min.js* from *dist* directory

2. Add the following javascript include tags to your html file:
    ```html
    <script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/jquery/1.8.0/jquery.min.js"></script>
    <script type="text/javascript" src="exceljs.compiled.min.js"></script>
    ```
3. Add the following javascript to your html head section:
    ```html
    <script type="text/javascript">
        $(document).ready(function (data) {
            
            $("#downloader").click(function(){
                ExcelJS.createHtmlTableExcelFile('#myTable', 'excel-js-demo-export.xlsx');
			      });
		    });
    </script>
    ```
4. Create a HTML table like this
    
    ```html
    <table id="myTable">
        <tr><td>A1</td><td>B1</td><td>C1</td><td>D1</td></tr>
        <tr><td>A2</td><td>B2</td><td>C2</td><td>D2</td></tr>
        <tr><td>A3</td><td>B3</td><td>C3</td><td>D3</td></tr>
    </table>
    ```
    
5. Add a download link for your excel export file:
    ```html
    <a href="#" id="downloader">Download Excel!</a>
    ```

You can find the complete getting started example in the demo directory.
