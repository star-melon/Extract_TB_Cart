#!/usr/bin/python

# 请把购物车在浏览器里全部加载好后在开发者工具network中XHR复制有相关内容的json文件，保存到本地

import json
import xlwt

input_file_name = './cartmaker1.json'
output_file_name = './output1.xls'


def extract(json_file_name, xls_file_name):
    json_file = open(json_file_name, 'r', encoding='UTF-8')
    excel = xlwt.Workbook()
    sheet = excel.add_sheet('tb_cart')

    headers = ['店铺', '名称', '型号', '单价', '数量', '总价', '链接']
    for col in range(len(headers)):
        sheet.write(0, col, headers[col])

    raw = 1
    t = json.load(json_file)

    shops = t['list']
    for shop in shops:

        shop_name = shop['title']
        bundles = shop['bundles']

        for bundle in bundles:
            orders = bundle['orders']
            sheet.write(raw, 0, shop_name)

            for order in orders:
                # 链接
                url = 'https:' + order['url']
                # 商品名称
                title = order['title']
                # 数量
                amount = order['amount']['now']
                # 单价(源单位：分)
                price = order['price']['now']/100
                # 总价(源单位：分)
                total = order['price']['sum']/100
                # 颜色分类
                if(order['skuStatus'] == 2):
                    skuId = order['skuId']
                    color = str(order['skus'])
                    url = url + '&skuId=' + skuId
                else:
                    color = '--'

                print('   -> ', title, amount, price, total, color)
                print('\t', url)
                sheet.write(raw, 1, title)
                sheet.write(raw, 2, color)
                sheet.write(raw, 3, price)
                sheet.write(raw, 4, amount)
                sheet.write(raw, 5, total)
                sheet.write(raw, 6, url)
                raw += 1

    excel.save(xls_file_name)
    json_file.close()


print("转换 \"", input_file_name, "\" ==> \"", output_file_name, "\"")
extract(input_file_name, output_file_name)