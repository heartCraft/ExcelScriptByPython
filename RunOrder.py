import sys
import xlwings as xw

# 打开excel文件
app=xw.App(visible=True,add_book=False)
app.display_alerts=False
app.screen_updating=False

# 打开三个excel文件获取数据
print('start!')
orderBook=app.books.open(r'C:\Users\Alisa\Desktop\配件交期自动化\runningOrder.xls')
orderSheet=orderBook.sheets[0]

storeBook=app.books.open(r'C:\Users\Alisa\Desktop\配件交期自动化\库存.xls')

deliveryBook=app.books.open(r'C:\Users\Alisa\Desktop\配件交期自动化\运输汇总.xlsx')

try:
    # 订单表数据读取
    orderTitle=orderSheet.range('a1').expand('right').value
    orderData=orderSheet.range('a1').expand('table').value[1:]
    if '缺货量' in orderTitle:
        index=orderTitle.index('缺货量')
        orderTitle=orderTitle[:index]
        orderData=[item[:index] for item in orderData]

    # 订单排序(销售需求交期>备注)
    orderData.sort(key=lambda item: (item[orderTitle.index('销售需求交期')], item[orderTitle.index('备注')]))

    # 库存表数据读取
    storeTable=storeBook.sheets[0].range('a1').expand('table').value
    storeKeyIndex=storeTable[0].index('物料代码')
    storeNumIndex=storeTable[0].index('库存')

    # 物流汇总表数据读取
    deliveryTable=deliveryBook.sheets[0].range('a1').expand('table').value
    deliveryKeyIndex=deliveryTable[0].index('件号')
    deliveryCount=len(deliveryTable[0][deliveryKeyIndex+1:])
    
    # 标题append
    orderTitle.append('缺货量')
    orderTitle.append('库存分配')
    orderTitle.extend(deliveryTable[0][deliveryKeyIndex+1:])
    orderTitle.append('需采购数量')
    
    
    # 总缺货量，库存分配，各批运输分配
    keyIndex = orderTitle.index('产品名称')
    longKeyIndex = orderTitle.index('产品长代码')
    numIndex = orderTitle.index('数量')
    doneNumIndex = orderTitle.index('已发货数量')
    for item in orderData:
        need=item[numIndex]-item[doneNumIndex]
        # 总缺货量
        if need>0:
            item.append(need)
        else:
            item.append(None)
        
        # 库存分配
        storeMatch=[store for store in storeTable if store[storeKeyIndex]==item[longKeyIndex]]
        if len(storeMatch) and storeMatch[0][storeNumIndex]:
            storeMatch=storeMatch[0]
            if need>storeMatch[storeNumIndex]:
                item.append(storeMatch[storeNumIndex])
                need-=storeMatch[storeNumIndex]
                storeMatch[storeNumIndex]=0
            else:
                item.append(need)
                storeMatch[storeNumIndex]-=need
                need=0
        else:
            item.append(None)
        
        # 物流汇总分配
        deliveryMatch=[delivery for delivery in deliveryTable if delivery[deliveryKeyIndex]==item[keyIndex]]
        if len(deliveryMatch):
            deliveryMatch=deliveryMatch[0]
            for i in range(deliveryKeyIndex+1, deliveryKeyIndex+1+deliveryCount):
                if need and deliveryMatch[i]:
                    if need>deliveryMatch[i]:
                        item.append(deliveryMatch[i])
                        need-=deliveryMatch[i]
                        deliveryMatch[i]=0
                    else:
                        item.append(need)
                        deliveryMatch[i]-=need
                        need=0
                else:
                    item.append(None)
        else:
            item.extend([None]*deliveryCount)

        # 最终缺货量        
        if need>0:
            item.append(need)
        else:
            item.append(None)

    # 保存
    print('saving!!!')
    orderSheet.range('a1').expand('table').clear()
    orderSheet.range('a1').expand('right').value=orderTitle
    orderSheet.range('a2').expand('table').value=orderData
    orderBook.save()
finally:
    orderBook.close()
    storeBook.close()
    deliveryBook.close()
    app.quit()
    print('ok!')





