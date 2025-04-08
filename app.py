from flask import Flask, render_template, request, jsonify, send_file
from sqlalchemy import create_engine, text
from urllib.parse import quote_plus
import pandas as pd
from datetime import datetime
import os
from dotenv import load_dotenv

app = Flask(__name__)
load_dotenv()

# Password for DB access
password = quote_plus("29031982")

# Database connection strings for each company
DATABASES = {
    "Hussain Traders": f"mssql+pyodbc://sa:{password}@192.168.10.2:1433/pharma_solution?driver=ODBC+Driver+17+for+SQL+Server",
    "Pharma Solution": f"mssql+pyodbc://sa:{password}@192.168.10.2:1433/PS_Trade?driver=ODBC+Driver+17+for+SQL+Server"
}

# Function to get the correct engine based on the company selected
def get_engine(company_name):
    conn_str = DATABASES.get(company_name)
    if not conn_str:
        raise ValueError(f"Unknown company: {company_name}")
    return create_engine(conn_str)

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/get_product_range')
def get_product_range():
    try:
        company_name = request.args.get("company_name")
        
        if not company_name:
            return jsonify({"success": False, "error": "Missing company name"})
            
        engine = get_engine(company_name)

        with engine.connect() as conn:
            # Updated queries to get correct first and last product
            first_query = text("""
                SELECT TOP 1 
                    productId, 
                    ProductPack
                FROM Product 
                ORDER BY productId asc
            """)
            
            last_query = text("""
                SELECT TOP 1 
                    productId, 
                    ProductPack
                FROM Product 
                ORDER BY productId DESC
            """)

            first = conn.execute(first_query).fetchone()
            last = conn.execute(last_query).fetchone()

            if not first or not last:
                return jsonify({
                    "success": False,
                    "error": "No products found in the database"
                })

        return jsonify({
            "success": True,
            "first_id": first[0] if first else None,
            "first_name": f"{first[1]}" if first else None,
            "last_id": last[0] if last else None,
            "last_name": f"{last[1]}" if last else None
        })

    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return jsonify({"success": False, "error": str(e)})
    

@app.route('/get_suppliers')
def get_suppliers():
    try:
        company_name = request.args.get("company_name")  # Get selected company from the request
        engine = get_engine(company_name)  # Dynamically choose the database based on company

        with engine.connect() as conn:
            # Query to get unique suppliers based on company
            query = """
            SELECT acc4, TITLE 
            FROM (SELECT * FROM V_M_PARTY) AS TR7  
            WHERE ACC3 = '3-03-11' AND Id LIKE '%3-03-11-%' 
            ORDER BY ACC4
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
        selected_suppliers = data.get('suppliers', [])
        company_name = data.get('company_name')
        start_product = data.get('start_product', '0001')
        end_product = data.get('end_product', '8600')
        report_format = data.get('report_format', 'IBL')  # Default to IBL format

        if not all([start_date, end_date, company_name, selected_suppliers]):
            return jsonify({"success": False, "error": "Missing required parameters"})

        company_codes = {
            "Hussain Traders": {
                "RD_code": "9200000006",
                "IBL_Branch_code": "9206",
                "Franchise_Code": "9200000006"
            },
            "Pharma Solution": {
                "RD_code": "9200000007",
                "IBL_Branch_code": "9207",
                "Franchise_Code": "9200000007"
            }
        }
        
        codes = company_codes.get(company_name)
        if not codes:
            return jsonify({"success": False, "error": "Invalid company selected"})

        engine = get_engine(company_name)
        
        # Create filename based on format and company
        if report_format == "Hudson":
            file_name = f"Hudson_{company_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        else:  # IBL format
            file_name = f"IBLHC_{codes['RD_code']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            
        file_path = os.path.join(os.getcwd(), file_name)

        # Replace suppliers in queries
        suppliers_str = f"({','.join([f"'{s}'" for s in selected_suppliers])})"

        if report_format == "Hudson":
            # Hudson format query
            query_hudson = f"""
            SELECT 
                REPLACE(CONVERT(CHAR(15), CAST(DOCDATE AS DATETIME), 105), '-', '/') AS date,
                ProductPack AS itemname,
                RIGHT(CAST(productPackId AS VARCHAR(11)), 7) AS code,
                PackDesc AS pack,
                CAST(ISNULL(Rate, 0) AS DECIMAL(12,2)) AS price,
                BATCHNO AS batchname,
                RIGHT('00' + CAST(DOCUMENTNO AS VARCHAR(11)), 7) AS invno,
                REPLACE(CUSTID, '-', '') AS partycode,
                Party,
                REPLACE(TownId, '-', '') AS areacode,
                Town AS areaname,
                'Karachi' AS cityname,
                'RESP & INJECT GROUP' AS groupname,
                Qty,
                CAST(ISNULL(Rate, 0) AS DECIMAL(12,2)) AS tpamt,
                ISNULL(
                    CAST(
                        (
                            CAST(ISNULL(Rate, 0) * (ISNULL(DISC, 0) / 100) * Qty AS DECIMAL(12,2)) * 100
                        ) / NULLIF(CAST(ISNULL(Rate, 0) * Qty AS DECIMAL(12,2)), 0)
                        AS DECIMAL(12,2)
                    ),
                    0.00
                ) AS disperc,
                CAST(
                    ISNULL(Rate, 0) * (ISNULL(DISC, 0) / 100) * Qty 
                    AS DECIMAL(12,2)
                ) AS DISCOUNTAMT,
                CAST(
                    (Rate * Qty) - ((Rate * (ISNULL(DISC, 0) / 100)) * Qty)
                    AS DECIMAL(12,2)
                ) AS NETAMNT,
                REASON
            FROM 
                [dbo].[PAP_SI_ALL](:start_date, :end_date, N'SR', 'DCO') AS TR
            WHERE  
                LEFT(productPackId, 4) BETWEEN :start_product AND :end_product
                AND SupplierId IN :suppliers
            ORDER BY 
                DocType,
                DOCUMENTNO
            """
            
            query_hudson = query_hudson.replace(":suppliers", suppliers_str)
            
            with engine.connect() as conn:
                try:
                    result_hudson = conn.execute(
                        text(query_hudson),
                        {
                            "start_date": start_date,
                            "end_date": end_date,
                            "start_product": start_product,
                            "end_product": end_product
                        }
                    )
                    df_hudson = pd.DataFrame(result_hudson.fetchall(), columns=result_hudson.keys())
                    
                    # Create Excel file with Hudson format
                    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                        if not df_hudson.empty:
                            df_hudson.to_excel(writer, sheet_name='Sales', index=False)
                        else:
                            pd.DataFrame({'Message': ['No sales data available for the selected criteria']}).to_excel(
                                writer, sheet_name='Sales', index=False
                            )
                    
                    return send_file(
                        file_path,
                        as_attachment=True,
                        download_name=file_name,
                        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                except Exception as e:
                    print(f"Error fetching Hudson data: {e}")
                    return jsonify({"success": False, "error": "Error fetching Hudson data"})
        else:
            # IBL format queries (existing code)
            query_sales = f"""
                SELECT 
                    RIGHT('00'+ISNULL(CAST(DOCUMENTNO AS VARCHAR(11)),''),7) AS Franchise_Customer_OrderNo,
                    REPLACE(CONVERT(CHAR(10), CAST(DOCDATE AS DATETIME), 101), '-', '/') Franchise_Customer_Invoice_Date,
                    RIGHT('00'+ISNULL(CAST(DOCUMENTNO AS VARCHAR(11)),''),7) AS Franchise_Customer_Invoice_Number,
                    RIGHT(ISNULL(CustType, ''), CASE WHEN LEN(ISNULL(CustType, '')) > 3 THEN LEN(CustType) - 3 ELSE 0 END) AS Channel,
                    '{codes["Franchise_Code"]}' as Franchise_Code,
                    REPLACE(CUSTID, '-', '') Franchise_Customer_Number,
                    REPLACE(CUSTID, '-', '') IBL_Customer_Number,
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
                    replace(TownId, '-', '') as Brick_Code,
                    Town as Brick_Name
                FROM [dbo].[PAP_SI_ALL](:start_date, :end_date, N'SR', 'DCO') AS TR
                WHERE LEFT(productPackId,4) BETWEEN :start_product AND :end_product
                AND SupplierId IN :suppliers
                AND ProductPack not like '%f.o.c %'
                ORDER BY DocType, DOCUMENTNO
                """
            
            query_stocks = f"""
                SELECT 
                    '{codes["RD_code"]}' AS RD_code,
                    '{codes["IBL_Branch_code"]}' AS IBL_Branch_code,
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
                    FROM UDF_PAM_OPG_IBL(:end_date, :start_product, :end_product, 0) 
                    GROUP BY LEFT(productId, 4), packId, batchNo, expiry
                ) OPG
                INNER JOIN Product 
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
            
            query_sales = query_sales.replace(":suppliers", suppliers_str)
            query_stocks = query_stocks.replace(":suppliers", suppliers_str)

            with engine.connect() as conn:
                # Execute queries with proper error handling
                try:
                    result_sales = conn.execute(
                        text(query_sales),
                        {
                            "start_date": start_date,
                            "end_date": end_date,
                            "start_product": start_product,
                            "end_product": end_product
                        }
                    )
                    df_sales = pd.DataFrame(result_sales.fetchall(), columns=result_sales.keys())
                    df_sales.columns = [col.replace("_", " ") for col in df_sales.columns]
                except Exception as e:
                    print(f"Error fetching sales data: {e}")
                    df_sales = pd.DataFrame()

                try:
                    result_stocks = conn.execute(
                        text(query_stocks),
                        {
                            "end_date": end_date,
                            "start_product": start_product,
                            "end_product": end_product
                        }
                    )
                    df_stocks = pd.DataFrame(result_stocks.fetchall(), columns=result_stocks.keys())
                    df_stocks.columns = [col.replace("_", " ") for col in df_stocks.columns]
                except Exception as e:
                    print(f"Error fetching stock data: {e}")
                    df_stocks = pd.DataFrame()

                # Create Excel writer with proper error handling
                try:
                    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                        if not df_sales.empty:
                            df_sales.to_excel(writer, sheet_name='Sales', index=False)
                        else:
                            pd.DataFrame({'Message': ['No sales data available for the selected criteria']}).to_excel(
                                writer, sheet_name='Sales', index=False
                            )

                        if not df_stocks.empty:
                            df_stocks.to_excel(writer, sheet_name='Stocks', index=False)
                        else:
                            pd.DataFrame({'Message': ['No stock data available for the selected criteria']}).to_excel(
                                writer, sheet_name='Stocks', index=False
                            )

                    return send_file(
                        file_path,
                        as_attachment=True,
                        download_name=file_name,
                        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                except Exception as e:
                    print(f"Error creating Excel file: {e}")
                    return jsonify({"success": False, "error": "Error creating Excel file"})

    except Exception as e:
        print(f"General error: {e}")
        return jsonify({"success": False, "error": str(e)})

if __name__ == '__main__':
    app.run(debug=True , host='0.0.0.0',port=5000)
