from templates.dash_app import app , server, submissions, handle_all_inputs

if __name__=="__main__":
    app.run(debug=True) 
    print(f'submission::{submissions}')
    print(f"email id{handle_all_inputs.get('email')}")

app.index_string = '''
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>{%title%}</title>
        {%favicon%}
        {%css%}
        <style>
            * {
                margin: 0 !important;
                padding: 0 !important;
                box-sizing: border-box !important;
            }
            html, body {
                height: 100% !important;
                overflow-x: hidden !important;
                margin: 0 !important;
                padding: 0 !important;
            }
            #react-entry-point {
                height: 100vh !important;
                margin: 0 !important;
                padding: 0 !important;
            }
            .form-field, .Select-control, .Select--single, .Select-placeholder, .Select-value, .Select-input, .Select-menu-outer, .Select-menu {
                min-height: 40px !important;
                height: 40px !important;
                font-size: 1rem !important;
                border-radius: 8px !important;
                border: 1px solid #ced4da !important;
                background: #fff !important;
                box-shadow: 0 1px 4px rgba(0,0,0,0.04);
                padding-left: 12px !important;
                padding-right: 12px !important;
                box-sizing: border-box;
                display: flex;
                align-items: center;
            }
            input.form-field.form-control {
                line-height: 2 !important;
                height: 40px !important;
            }
            .Select-control {
                border: 1px solid #ced4da !important;
                background: #fff !important;
                box-shadow: 0 1px 4px rgba(0,0,0,0.04);
                padding-left: 12px !important;
                padding-right: 12px !important;
                min-height: 40px !important;
                height: 40px !important;
                border-radius: 8px !important;
                display: flex;
                align-items: center;
            }
            .Select-placeholder, .Select-value {
                font-size: 1rem !important;
                color: #6c757d !important;
                line-height: 2.0 !important;
            }
            .Select-arrow-zone {
                padding-top: 0 !important;
                padding-bottom: 0 !important;
                display: flex;
                align-items: center;
            }
            .Select-menu-outer {
                border-radius: 8px !important;
                font-size: 1rem !important;
            }
            .form-field:focus, .Select-control:focus {
                border-color: #2684ff !important;
                box-shadow: 0 0 0 2px rgba(38,132,255,0.2);
            }
        </style>
    </head>
    <body>
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
    </body>
</html>
'''