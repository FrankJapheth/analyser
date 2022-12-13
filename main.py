import requests
from dalali_products import DalaliProducts, format_property, find_values, get_redundancy

if __name__ == '__main__':
    dalali_products = DalaliProducts()

    dalali_products.read_products_file(path_to_file='files/latest_prices.xlsx')
    sheet_data = dalali_products.get_products_details(sheet_name='Table 1', max_col=2, max_row=1809)

    list_a = []

    list_b = []

    for product_details in sheet_data['data']:
        list_a.append(
            format_property(
                product_details['Product Name'],
                ['rmvwp', 'rmvftp', 'mksml', 'rmvnl']
            )
        )

    dalali_products.read_products_file(path_to_file='files/db_products.xlsx')
    sheet_data2 = dalali_products.get_products_details(sheet_name='products', max_col=10, max_row=1800)

    for product_details in sheet_data2['data']:
        list_b.append(
            format_property(product_details['name'] + product_details['metrics'],
                            ['rmvwp', 'rmvftp', 'mksml']
                            )
        )

    # get_redundancy(list_a)
    props_data = find_values(list_b, list_a)
    '''sheet_data3 = []
    for prop_not_present in props_data['properties_not_present']:
        sheet_data3.append(sheet_data2['data'][prop_not_present['index_in_a']])

    wb_sheet = DalaliProducts()
    wb_sheet.create_workbook()
    wb_sheet.write_data(sheet_data3, 'files/prods_n_found.xlsx')
    '''

    for data in props_data['properties_present']:
        sheet_data_prod_props = sheet_data['data'][data['index_in_b']]
        sheet_data2_prod_props = sheet_data2['data'][data['index_in_a']]

        product_to_update = {
            'productId': sheet_data2_prod_props['db_index'],
            'productCategoryid': sheet_data2_prod_props['cat_id'],
            'productName': sheet_data2_prod_props['name'],
            'productMetrics': sheet_data2_prod_props['metrics'],
            'status': sheet_data2_prod_props['status'],
            'productOUId': sheet_data2_prod_props['prod_order'],
            'prodImg': sheet_data2_prod_props['media'],
            'productPrice': sheet_data_prod_props['Price'],
            'wholesaleProductPrice': sheet_data_prod_props['Price'],
            'productQuantity': 50
        }

        url = 'https://dalaliwinehouse.com/backend/changeProd'

        r = requests.post(url, json=product_to_update, verify=False)

        print(r.json())
