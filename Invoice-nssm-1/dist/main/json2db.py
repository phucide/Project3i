import pyodbc
import json
import xml.etree.ElementTree as ET
xml_file = 'Invoice_Setting.xml'
tree = ET.parse(xml_file)
root = tree.getroot()




variable = root.find("variable")
db_location = variable.find("db_location")

print("Insert data to database: Buy In Invoice")
try:
    
    con_string = (
        r'DRIVER= {Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+db_location.text
    )    
    conn = pyodbc.connect(con_string)
    print("Connect to database")
    # Read the JSON file with UTF-8 encoding
    with open(variable.find("json_file").text, 'r', encoding='utf-8') as file:
        data = json.load(file)

    invoices_list = data["invoices"]
    if invoices_list:
        #print(type(invoices_list))
        invoice_count=1
        for invoice in invoices_list:
            print("Insert Invoice:{}".format(invoice_count))
            cursor = conn.cursor()
            
            ###Invoices header
            insert_query = """
            INSERT INTO Invoices_header
            (Inv_form, Inv_sign, Inv_num, Inv_title, Inv_date_express, inv_code, cus_name, cus_tax, cus_address, cus_tel, cus_bank_acc, buyer_name, full_name, buyer_tax, buyer_addr, buyer_account, payment_method,unit,list_number
            ,date_list, Buy_In_Status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ? , ?, ?, ?, ?)
            """
            values = (invoice["Mẫu_số"], invoice["Ký hiệu"], invoice["Số"], invoice["Tên hóa đơn"],
                    invoice["Ngày thành lập"], invoice["MCCQT:"], invoice["Tên người bán:"], invoice["Mã số thuế:"],
                    invoice["Địa chỉ:"], invoice["Điện thoại:"], invoice["Số tài khoản:"], invoice["Tên người mua:"],invoice["Họ tên người mua:"], invoice["Mã số thuế: buyer"],invoice["Địa chỉ: buyer"],invoice["Số tài khoản: buyer"],
                    invoice["Hình thức thanh toán:"],invoice["Đơn vị tiền tệ:"],invoice["Số bảng kê:"],invoice["Ngày bảng kê:"], 1)
            
            cursor.execute(insert_query, values)
            conn.commit()
            
            
            #####Invoice_item
            table_list = invoice["tables"]
            table_1 = table_list[0]
            rows = table_1["rows"]
            for row in rows:
                
                insert_query = """
                INSERT INTO Invoice_Item
                (Good_property,Good_name,Unit,Amount,Price,Discount,Tax,GTGT,Pay,Invoice_code,No)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ? , ?)
                """
                if "Thuế suất" in row:
                    tax = row["Thuế suất"]
                    gtgt = row["Thành tiền chưa có thuế GTGT"]
                    pay = "null"
                else:
                    tax = "null"
                    gtgt = "null"
                    pay = row["Thành tiền"]

               
                values = (
                    row["Tính chất"],row["Tên hàng hóa, dịch vụ"],row["Đơn vị tính"],row["Số lượng"],
                    row["Đơn giá"],row["Chiết khấu"],tax,gtgt,pay,
                    invoice["MCCQT:"],row["STT"]
                )
                cursor.execute(insert_query, values)
                conn.commit()
                
            
            
            table_2 = table_list[1]
            table_3 = table_list[2]
            table_4 = table_list[3]
            
            ###table_4:
            if "Tổng tiền chưa thuế\n(Tổng cộng thành tiền chưa có thuế)" in table_4:
                    # If the key exists, assign its value to item1
                item1 = table_4["Tổng tiền chưa thuế\n(Tổng cộng thành tiền chưa có thuế)"]
            else:
                # If the key doesn't exist, assign "null" to item1
                item1 = "null"
            
            if "Tổng giảm trừ không chịu thuế" in table_4:
                    # If the key exists, assign its value to item1
                item2 = table_4["Tổng giảm trừ không chịu thuế"]
            else:
                # If the key doesn't exist, assign "null" to item1
                item2 = "null"
            
            if "Tổng tiền thuế (Tổng cộng tiền thuế)" in table_4:
                    # If the key exists, assign its value to item1
                item3 = table_4["Tổng tiền thuế (Tổng cộng tiền thuế)"]
            else:
                # If the key doesn't exist, assign "null" to item1
                item3 = "null"

            if "Tổng tiền phí" in table_4:
                    # If the key exists, assign its value to item1
                item4 = table_4["Tổng tiền phí"]
            else:
                # If the key doesn't exist, assign "null" to item1
                item4 = "null"

            if "Tổng tiền chiết khấu thương mại" in table_4:
                    # If the key exists, assign its value to item1
                item5 = table_4["Tổng tiền chiết khấu thương mại"]
            else:
                # If the key doesn't exist, assign "null" to item1
                item5 = "null"

            if "Tổng giảm trừ khác" in table_4:
                    # If the key exists, assign its value to item1
                item6 = table_4["Tổng giảm trừ khác"]
            else:
                # If the key doesn't exist, assign "null" to item1
                item6 = "null"

            if "Tổng tiền thanh toán bằng số" in table_4:
                    # If the key exists, assign its value to item1
                item7 = table_4["Tổng tiền thanh toán bằng số"]
            else:
                # If the key doesn't exist, assign "null" to item1
                item7 = "null"

            if "Tổng tiền thanh toán bằng chữ" in table_4:
                    # If the key exists, assign its value to item1
                item8 = table_4["Tổng tiền thanh toán bằng chữ"]
            else:
                # If the key doesn't exist, assign "null" to item1
                item8 = "null"
            
            if "Ghi chú" in table_4:
                    # If the key exists, assign its value to item1
                item9 = table_4["Ghi chú"]
            else:
                # If the key doesn't exist, assign "null" to item1
                item9 = "null"

            insert_query = """
                INSERT INTO Payment_Sum
                (item1,item2,item3,item4,item5,item6,item7,item8,item9,Inv_code)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """
            values = (
                    item1,item2,item3,item4,item5,item6,item7,item8,item9, invoice["MCCQT:"]
                )
            cursor.execute(insert_query, values)
            conn.commit()

            ##table 2
            if table_2["rows"]:
                if "Thuế suất" in table_2["rows"][0]:
                    for row in table_2["rows"]:
                        insert_query = """
                        INSERT INTO Tax
                        (Tax,Total_no_tax,Total_tax,Inv_code)
                        VALUES (?, ?, ?, ?)
                        """
                        values = (
                                row["Thuế suất"],row["Tổng tiền chưa thuế"],row["Tổng tiền thuế"], invoice["MCCQT:"]
                            )
                        cursor.execute(insert_query, values)
                        conn.commit()
                if "Tên loại phí" in table_2["rows"][0]:
                    for row in table_2["rows"]:
                        insert_query = """
                        INSERT INTO Fee
                        (No,Fee_name,Fee,Inv_code)
                        VALUES (?, ?, ?, ?)
                        """
                        values = (
                                row["STT"],row["Tên loại phí"],row["Tiền phí"], invoice["MCCQT:"]
                            )
                        cursor.execute(insert_query, values)
                        conn.commit()
                
             ##table 3
            if table_3["rows"]:
                if "Thuế suất" in table_3["rows"][0]:
                    for row in table_3["rows"]:
                        insert_query = """
                        INSERT INTO Tax
                        (Tax,Total_no_tax,Total_tax,Inv_code)
                        VALUES (?, ?, ?, ?)
                        """
                        values = (
                                row["Thuế suất"],row["Tổng tiền chưa thuế"],row["Tổng tiền thuế"], invoice["MCCQT:"]
                            )
                        cursor.execute(insert_query, values)
                        conn.commit()
                if "Tên loại phí" in table_3["rows"][0]:
                    for row in table_3["rows"]:
                        insert_query = """
                        INSERT INTO Fee
                        (No,Fee_name,Fee,Inv_code)
                        VALUES (?, ?, ?, ?)
                        """
                        values = (
                                row["STT"],row["Tên loại phí"],row["Tiền phí"], invoice["MCCQT:"]
                            )
                        cursor.execute(insert_query, values)
                        conn.commit()



            invoice_count+=1
        cursor.close()
        conn.close()
        print("Sucessful insert all invoices")
        print("Process completely successfull")
    else:
        print("No invoices added")
except pyodbc.Error as e:
    print('Error',e)


print("________________________")
print("________________________")
print("Insert data to database: Sell Out Invoice")
try:
    
    con_string = (
        r'DRIVER= {Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+db_location.text
    )    
    conn = pyodbc.connect(con_string)
    print("Connect to database")
    # Read the JSON file with UTF-8 encoding
    with open(variable.find("json_out").text, 'r', encoding='utf-8') as file:
        data = json.load(file)

    invoices_list = data["invoices"]
    if invoices_list:
        #print(type(invoices_list))
        invoice_count=1
        for invoice in invoices_list:
            print("Insert Invoice:{}".format(invoice_count))
            cursor = conn.cursor()
            
            ###Invoices header
            insert_query = """
            INSERT INTO Invoices_header
            (Inv_form, Inv_sign, Inv_num, Inv_title, Inv_date_express, inv_code, cus_name, cus_tax, cus_address, cus_tel, cus_bank_acc, buyer_name, full_name, buyer_tax, buyer_addr, buyer_account, payment_method,unit,list_number
            ,date_list, Buy_In_Status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ? , ?, ?, ?, ?)
            """
            values = (invoice["Mẫu_số"], invoice["Ký hiệu"], invoice["Số"], invoice["Tên hóa đơn"],
                    invoice["Ngày thành lập"], invoice["MCCQT:"], invoice["Tên người bán:"], invoice["Mã số thuế:"],
                    invoice["Địa chỉ:"], invoice["Điện thoại:"], invoice["Số tài khoản:"], invoice["Tên người mua:"],invoice["Họ tên người mua:"], invoice["Mã số thuế: buyer"],invoice["Địa chỉ: buyer"],invoice["Số tài khoản: buyer"],
                    invoice["Hình thức thanh toán:"],invoice["Đơn vị tiền tệ:"],invoice["Số bảng kê:"],invoice["Ngày bảng kê:"], 0)
            
            cursor.execute(insert_query, values)
            conn.commit()
            
            
            #####Invoice_item
            table_list = invoice["tables"]
            table_1 = table_list[0]
            rows = table_1["rows"]
            for row in rows:
                
                insert_query = """
                INSERT INTO Invoice_Item
                (Good_property,Good_name,Unit,Amount,Price,Discount,Tax,GTGT,Pay,Invoice_code,No)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ? , ?)
                """
                if "Thuế suất" in row:
                    tax = row["Thuế suất"]
                    gtgt = row["Thành tiền chưa có thuế GTGT"]
                    pay = "null"
                else:
                    tax = "null"
                    gtgt = "null"
                    pay = row["Thành tiền"]

               
                values = (
                    row["Tính chất"],row["Tên hàng hóa, dịch vụ"],row["Đơn vị tính"],row["Số lượng"],
                    row["Đơn giá"],row["Chiết khấu"],tax,gtgt,pay,
                    invoice["MCCQT:"],row["STT"]
                )
                cursor.execute(insert_query, values)
                conn.commit()
                
            
            
            table_2 = table_list[1]
            table_3 = table_list[2]
            table_4 = table_list[3]
            
            ###table_4:
            if "Tổng tiền chưa thuế\n(Tổng cộng thành tiền chưa có thuế)" in table_4:
                    # If the key exists, assign its value to item1
                item1 = table_4["Tổng tiền chưa thuế\n(Tổng cộng thành tiền chưa có thuế)"]
            else:
                # If the key doesn't exist, assign "null" to item1
                item1 = "null"
            
            if "Tổng giảm trừ không chịu thuế" in table_4:
                    # If the key exists, assign its value to item1
                item2 = table_4["Tổng giảm trừ không chịu thuế"]
            else:
                # If the key doesn't exist, assign "null" to item1
                item2 = "null"
            
            if "Tổng tiền thuế (Tổng cộng tiền thuế)" in table_4:
                    # If the key exists, assign its value to item1
                item3 = table_4["Tổng tiền thuế (Tổng cộng tiền thuế)"]
            else:
                # If the key doesn't exist, assign "null" to item1
                item3 = "null"

            if "Tổng tiền phí" in table_4:
                    # If the key exists, assign its value to item1
                item4 = table_4["Tổng tiền phí"]
            else:
                # If the key doesn't exist, assign "null" to item1
                item4 = "null"

            if "Tổng tiền chiết khấu thương mại" in table_4:
                    # If the key exists, assign its value to item1
                item5 = table_4["Tổng tiền chiết khấu thương mại"]
            else:
                # If the key doesn't exist, assign "null" to item1
                item5 = "null"

            if "Tổng giảm trừ khác" in table_4:
                    # If the key exists, assign its value to item1
                item6 = table_4["Tổng giảm trừ khác"]
            else:
                # If the key doesn't exist, assign "null" to item1
                item6 = "null"

            if "Tổng tiền thanh toán bằng số" in table_4:
                    # If the key exists, assign its value to item1
                item7 = table_4["Tổng tiền thanh toán bằng số"]
            else:
                # If the key doesn't exist, assign "null" to item1
                item7 = "null"

            if "Tổng tiền thanh toán bằng chữ" in table_4:
                    # If the key exists, assign its value to item1
                item8 = table_4["Tổng tiền thanh toán bằng chữ"]
            else:
                # If the key doesn't exist, assign "null" to item1
                item8 = "null"
            
            if "Ghi chú" in table_4:
                    # If the key exists, assign its value to item1
                item9 = table_4["Ghi chú"]
            else:
                # If the key doesn't exist, assign "null" to item1
                item9 = "null"

            insert_query = """
                INSERT INTO Payment_Sum
                (item1,item2,item3,item4,item5,item6,item7,item8,item9,Inv_code)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """
            values = (
                    item1,item2,item3,item4,item5,item6,item7,item8,item9, invoice["MCCQT:"]
                )
            cursor.execute(insert_query, values)
            conn.commit()

            ##table 2
            
            if table_2["rows"]:
                if "Thuế suất" in table_2["rows"][0]:
                    for row in table_2["rows"]:
                        insert_query = """
                        INSERT INTO Tax
                        (Tax,Total_no_tax,Total_tax,Inv_code)
                        VALUES (?, ?, ?, ?)
                        """
                        values = (
                                row["Thuế suất"],row["Tổng tiền chưa thuế"],row["Tổng tiền thuế"], invoice["MCCQT:"]
                            )
                        cursor.execute(insert_query, values)
                        conn.commit()
                if "Tên loại phí" in table_2["rows"][0]:
                    for row in table_2["rows"]:
                        insert_query = """
                        INSERT INTO Fee
                        (No,Fee_name,Fee,Inv_code)
                        VALUES (?, ?, ?, ?)
                        """
                        values = (
                                row["STT"],row["Tên loại phí"],row["Tiền phí"], invoice["MCCQT:"]
                            )
                        cursor.execute(insert_query, values)
                        conn.commit()
                
             ##table 3
            if table_3["rows"]:
                if "Thuế suất" in table_3["rows"][0]:
                    for row in table_3["rows"]:
                        insert_query = """
                        INSERT INTO Tax
                        (Tax,Total_no_tax,Total_tax,Inv_code)
                        VALUES (?, ?, ?, ?)
                        """
                        values = (
                                row["Thuế suất"],row["Tổng tiền chưa thuế"],row["Tổng tiền thuế"], invoice["MCCQT:"]
                            )
                        cursor.execute(insert_query, values)
                        conn.commit()
                if "Tên loại phí" in table_3["rows"][0]:
                    for row in table_3["rows"]:
                        insert_query = """
                        INSERT INTO Fee
                        (No,Fee_name,Fee,Inv_code)
                        VALUES (?, ?, ?, ?)
                        """
                        values = (
                                row["STT"],row["Tên loại phí"],row["Tiền phí"], invoice["MCCQT:"]
                            )
                        cursor.execute(insert_query, values)
                        conn.commit()

            invoice_count+=1
        cursor.close()
        conn.close()
        print("Sucessful insert all invoices")
        print("Process completely successfull")
    else:
        print("No invoices added")
except pyodbc.Error as e:
    print('Error',e)