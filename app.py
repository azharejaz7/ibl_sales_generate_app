from flask import Flask, render_template, request, jsonify, send_file
from sqlalchemy import create_engine, text
from urllib.parse import quote_plus
import pandas as pd
from datetime import datetime
import os
from dotenv import load_dotenv

app = Flask(__name__)
load_dotenv()
password = quote_plus("29031982")


# Database connection
connection_string = f"mssql+pyodbc://sa:{password}@192.168.10.2:1433/PS_Trade?driver=ODBC+Driver+17+for+SQL+Server"
engine = create_engine(connection_string)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/get_suppliers')
def get_suppliers():
    try:
        with engine.connect() as conn:
            # Query to get unique suppliers
            query = """
            SELECT  acc4,TITLE FROM (SELECT * FROM V_M_PARTY)AS TR7  WHERE ACC3='3-03-11' AND Id like '%3-03-11-%' ORDER BY ACC4
            """
            result = conn.execute(text(query))
            suppliers = [{"id": row[0], "name": row[1]} for row in result]
            return jsonify({"success": True, "suppliers": suppliers})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})

@app.route('/get_data', methods=['POST'])
def get_data():
    try:
        data = request.json
        start_date = data.get('start_date')
        end_date = data.get('end_date')
        company_name = data.get('company_name', 'Hussain Traders')
        selected_suppliers = data.get('suppliers', [])
        
        # Define company codes based on company name
        company_codes = {
            'RD_code': '9200000006',
            'IBL_Branch_code': '9206',
            'Franchise_Code': '9200000006'
        }
        
        if company_name == 'Pharma Solution':
            company_codes = {
                'RD_code': '9200000007',
                'IBL_Branch_code': '9207',
                'Franchise_Code': '9200000007'
            }
        
        # Create a fixed path for the Excel file
        file_path = os.path.join(os.getcwd(), f"export_{company_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        
        # Create Excel writer
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            
            # Always generate both sheets regardless of form_type selection
            # First, generate Sales data (Form 1)
            query_sales = f"""
            SELECT RIGHT('00'+CAST(DOCUMENTNO AS VARCHAR(11)),7) Franchise_Customer_OrderNo,
            REPLACE(CONVERT(CHAR(10), CAST(DOCDATE AS DATETIME), 101), '-', '/') Franchise_Customer_Invoice_Date,
            RIGHT('00'+CAST(DOCUMENTNO AS VARCHAR(11)),7) Franchise_Customer_Invoice_Number,
            right(CustType,len(CustType)-3 ) Channel,
            '{company_codes["Franchise_Code"]}' as Franchise_Code,
            REPLACE(CUSTID,'-','') Franchise_Customer_Number,
            REPLACE(CUSTID,'-','') IBL_Customer_Number,
            Party as RD_Customer_Name,
            Party as IBL_Customer_Name,
            Town as Customer_Address,
            manufacturerCode as Franchise_Item_Code,
            manufacturerCode as IBL_Item_Code,
            ProductPack as Franchise_Item_Description,
            productpack as IBL_Item_Description,
            Qty Quantity_Sold,
            Cast(((Rate * qty)-(Rate*(ISNULL(DISC,0)/100)) * qty) as decimal(12,2)) Gross_Amount,
            REASON,
            CASE WHEN ISNULL(Free,0)=0 THEN ISNULL(Bonus,0) ELSE ISNULL(Free,0) END FOC,
            BATCHNO,
            Cast(((Rate)-(Rate*(ISNULL(DISC,0)/100))) as decimal(12,2)) PRICE,
            0 as BON_QTY,
            cast((ISNULL(Rate,0)*(ISNULL(DISC,0)/100) * qty) as decimal(12,2)) DISCOUNTAMT,
            Cast(((Rate * qty)-(Rate*(ISNULL(DISC,0)/100)) * qty) as decimal(12,2)) NET_AMT,
            0 as DISCOUNTED_RATE,
            replace(TownId,'-','') as Brick_Code,
            Town as Brick_Name 
            FROM [dbo].[PAP_SI_ALL](:start_date, :end_date, N'SR','DCO') AS TR
            WHERE LEFT(productPackId,4) BETWEEN '0001' AND '8600'
            AND SupplierId IN :suppliers
            AND ProductPack not like '%f.o.c %'
            ORDER BY DocType,DOCUMENTNO
            """
            
            # Replace suppliers placeholder in query
            query_sales = query_sales.replace(":suppliers", f"({','.join([f"'{s}'" for s in selected_suppliers])})")
            
            with engine.connect() as conn:
                result_sales = conn.execute(
                    text(query_sales),
                    {
                        "start_date": start_date,
                        "end_date": end_date
                    }
                )
                rows_sales = result_sales.fetchall()
                columns_sales = result_sales.keys()
                
                # Create DataFrame for Sales
                df_sales = pd.DataFrame(rows_sales, columns=columns_sales)
                df_sales.columns = [col.replace("_", " ") for col in df_sales.columns]
                
                # Write Sales sheet
                df_sales.to_excel(writer, index=False, sheet_name='Sales')
            
            # Second, generate Stocks data (Form 2)
            query_stocks = f"""
            SELECT 
                '{company_codes["RD_code"]}' AS RD_code,
                '{company_codes["IBL_Branch_code"]}' AS IBL_Branch_code,
                Product.manufacturerCode AS RD_Item_Code,
                Product.manufacturerCode AS IBL_Item_Code,
                Product.Product AS RD_Item_Description,
                OPG.batchNo AS LOT_NUMBER,
                OPG.expiry AS Expiry_Date,
                (CAST(OPG.opg AS INT) + ISNULL(PAE_SITR_Temp.QtyIn, 0)) AS Closing_Quantity,
                CASE 
                    WHEN Product.staxType = 'NET' THEN (CAST(OPG.opg AS INT) + ISNULL(PAE_SITR_Temp.QtyIn, 0)) * (Product.rateNew * 1.18)
                    WHEN Product.staxType = 'MRP' THEN (CAST(OPG.opg AS INT) + ISNULL(PAE_SITR_Temp.QtyIn, 0)) * (Product.MRP * 0.18) + Product.rateNew 
                    ELSE (CAST(OPG.opg AS INT) + ISNULL(PAE_SITR_Temp.QtyIn, 0)) * Product.rateNew
                END AS Value,
                REPLACE(CONVERT(CHAR(15), CAST(:end_date AS DATETIME), 101), '-', '/') AS Date,
                CASE 
                    WHEN Product.staxType = 'NET' THEN (Product.rateNew * 1.18)
                    WHEN Product.staxType = 'MRP' THEN (Product.MRP * 0.18) + Product.rateNew 
                    ELSE Product.rateNew
                END AS price,
                0 AS In_Transit_stock,
                0 AS Purchase_Unit
            FROM (
                SELECT 
                    LEFT(productId, 4) AS productId,
                    packId,
                    batchNo,
                    expiry,
                    AVG(avgRate) AS avgRate,
                    SUM(opg) AS opg
                FROM UDF_PAM_OPG_IBL(:end_date, '0001', '8600', 0) 
                GROUP BY LEFT(productId, 4), packId, batchNo, expiry
            ) OPG
            LEFT JOIN Product 
                ON Product.productPackId = CAST(OPG.productId AS VARCHAR(6)) + CAST(OPG.packId AS VARCHAR(6)) 
            LEFT JOIN (
                SELECT 
                    LEFT(productId, 4) AS productId,
                    PackId,
                    UPPER(batchNo) AS batchNo,
                    SUM(QtyIn) AS QtyIn
                FROM PAE_SITR_TEMp 
                GROUP BY LEFT(productId, 4), PackId, UPPER(batchNo)
            ) PAE_SITR_TEMp 
                ON PAE_SITR_Temp.productId = LEFT(OPG.productId, 4) 
                AND PAE_SITR_Temp.PackId = OPG.packId 
                AND UPPER(OPG.batchNo) = UPPER(PAE_SITR_Temp.BatchNo) 
            WHERE (CAST(OPG.opg AS INT) + ISNULL(PAE_SITR_Temp.QtyIn, 0)) <> 0 
            AND SupplierId IN :suppliers 
            AND Product.Product not like '%f.o.c%'
            GROUP BY 
                Product.manufacturerCode,  
                Product.Product,
                OPG.batchNo, 
                OPG.expiry, 
                Product.staxType,  
                Product.rateNew,  
                Product.MRP,
                (CAST(OPG.opg AS INT) + ISNULL(PAE_SITR_Temp.QtyIn, 0))
            """
            
            # Replace suppliers placeholder in query
            query_stocks = query_stocks.replace(":suppliers", f"({','.join([f"'{s}'" for s in selected_suppliers])})")
            
            with engine.connect() as conn:
                result_stocks = conn.execute(
                    text(query_stocks),
                    {
                        "end_date": end_date
                    }
                )
                rows_stocks = result_stocks.fetchall()
                columns_stocks = result_stocks.keys()
                
                # Create DataFrame for Stocks
                df_stocks = pd.DataFrame(rows_stocks, columns=columns_stocks)

                # Remove underscores from column names
                df_stocks.columns = [col.replace("_", " ") for col in df_stocks.columns]
                
                # Write Stocks sheet
                df_stocks.to_excel(writer, index=False, sheet_name='Stocks')
        
        # Return the file
        return send_file(
            file_path,
            as_attachment=True,
            download_name=os.path.basename(file_path),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return jsonify({"success": False, "error": str(e)})

if __name__ == '__main__':
    app.run(debug=True) 