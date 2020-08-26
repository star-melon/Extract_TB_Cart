#请把购物车在浏览器里全部加载好后在开发者工具network中XHR复制有相关内容的json文件，保存到本地

import json
import xlwt

#将json文件放入同一文件夹中，替换FILE_NAME
FILE_NAME = './cartmaker2.json'


excel = xlwt.Workbook()
sheet = excel.add_sheet('tb_cart')

sheet.write(0, 0, '店铺')
sheet.write(0, 1, '名称')
sheet.write(0, 2, '型号')
sheet.write(0, 3, '单价')
sheet.write(0, 4, '数量')
sheet.write(0, 5, '总价')
sheet.write(0, 6, '链接')

raw = 1
json_file = open(FILE_NAME, 'r', encoding='UTF-8')
print(json_file)
t = json.load(json_file)

shops = t['list']
for shop in shops:

    orders = shop['bundles'][0]['orders']
    shop_name = shop['seller']
    sheet.write(raw, 0, shop_name)

    for order in orders:
        #链接
        url = 'https:' + order['url']
        #商品名称
        title = order['title']
        #数量
        amount = order['amount']['now']
        #单价(源单位：分)
        price = order['price']['now']/100
        #总价(源单位：分)
        total = order['price']['sum']/100
        #颜色分类
        if(order['skuStatus'] == 2):
            skuId = order['skuId']
            color = str(order['skus'])
            url = url + '&skuId=' + skuId
        else:
            color = '默认'

        print('list------>', url, title, amount, price, total, color)
        sheet.write(raw, 1, title)
        sheet.write(raw, 2, color)
        sheet.write(raw, 3, price)
        sheet.write(raw, 4, amount)
        sheet.write(raw, 5, total)
        sheet.write(raw, 6, url)
        raw += 1 


excel.save('./output.xls')
