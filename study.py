import pandas as pd

# 读取整个文件
# df = pd.read_excel('I:\学习笔记\Python\BOM\BSTA4111-MB 材料清单260320 V1.5.xlsx')
# print(df.head())

# # 读取指定工作表
# df = pd.read_excel('I:\学习笔记\Python\BOM\BSTA4111-MB 材料清单260320 V1.5.xlsx', sheet_name='材料清单')
# print(df)

# # 或按索引读取
# df = pd.read_excel('data.xlsx', sheet_name=0)
#
# # 读取指定列
excel1='I:\学习笔记\Python\BOM\BSTA4111-MB 材料清单260320 V1.5.xlsx'
excel2='I:\学习笔记\Python\BOM\BSTA4111-MB 材料清单260320 V1.5 - 副本.xlsx'
df1 = pd.read_excel(excel1, sheet_name='材料清单',usecols=['Quantity', 'Reference'])
df2 = pd.read_excel(excel2, sheet_name='材料清单',usecols=['Quantity', 'Reference'])
quantity_list1=df1['Quantity']
quantity_list2=df2['Quantity']
reference_list1=df1['Reference'].tolist()
reference_list2=df2['Reference'].tolist()
big_list1 = list()
big_list2 = list()

# print(reference1[0])
# print(type(reference1))
# print(reference2)


# enumerate：同时获取 索引 和 值
if len(reference_list1) == len(reference_list2):
    print("两个BOM行数相同")
    # print('BOM1')
    num=len(reference_list1)
    print('BOM共有'+str(num)+'行\n')
    # print('原BOM1')
    # print(reference_list1)
    print('BOM1共有',sum(quantity_list1),'个器件')
    # print('原BOM2')
    # print(reference_list2)
    print('BOM2共有', sum(quantity_list2), '个器件')
    print('\n')
    # print(reference_list2)

    if sum(quantity_list1) != sum(quantity_list2):
        print('BOM1和BOM2元器件总数量不同，请检查BOM')
        for i in range (len(reference_list1)):
            first_list1 = reference_list1[i].split(',')
            for j in range(len(first_list1)):
                second_list1 = first_list1[j].split(',')
                big_list1.extend(second_list1)
        for i2 in range(len(reference_list1)):
            first_list2 = reference_list2[i].split(',')
            # print(first_list)
            for j2 in range(len(first_list1)):
                second_list2 = first_list2[j].split(',')
                # print(second_list)
                big_list2.extend(second_list2)

        # print(big_list1)
        print('第一个BOM共有',len(big_list1),'个')
        if len(set(big_list1)) != len(big_list1):
            print("BOM1中有重复值！")
            duplicates1 = list({x for x in big_list1 if big_list1.count(x) > 1})
            print('BOM1中重复位号：',duplicates1)
        else:
            print("BOM1没有重复值")
        print('\n')

        # print(big_list2)
        print('第二个BOM共有', len(big_list2), '个')
        if len(set(big_list2)) != len(big_list2):
            print("BOM2中有重复值！")
            duplicates2 = list({x for x in big_list2 if big_list2.count(x) > 1})
            print('BOM1中重复位号：', duplicates2)
        else:
            print("BOM2没有重复值")
        print('\n')

        diff1 = [x for x in big_list1 if x not in big_list2]
        diff2 = [x for x in big_list2 if x not in big_list1]
        if len(diff1) != 0:
            print('在BOM1但是不在BOM2的位号')
            print(diff1,'共',len(diff1),'个')
            df = pd.DataFrame(diff1, columns=["在BOM1但是不在BOM2的位号"])  # columns 是表头
            df.to_excel("在BOM1但是不在BOM2的位号.xlsx", index=False)
        else:
            print('BOM1中的位号都在BOM2中')
        if len(diff2) != 0:
            print('在BOM2但是不在BOM1的位号')
            print(diff2,'共',len(diff2),'个')
            df = pd.DataFrame(diff2, columns=["在BOM2但是不在BOM1的位号"])  # columns 是表头
            df.to_excel("在BOM2但是不在BOM1的位号.xlsx", index=False)
        else:
            print('BOM2中的位号都在BOM1中')

else:
    print("两个BOM行数不同,请检查BOM")
    print('BOM1行数：',len(quantity_list1))
    print('BOM2行数：',len(quantity_list2))
    print('BOM1共有', sum(quantity_list1), '个器件')
    # print('原BOM2')
    # print(reference_list2)
    print('BOM2共有', sum(quantity_list2), '个器件')
    print('\n')
    # print(reference_list2)

    if sum(quantity_list1) != sum(quantity_list2):
        print('BOM1和BOM2元器件总数量不同，请检查BOM')
    else:
        for i in range(len(reference_list1)):
            first_list1 = reference_list1[i].split(',')
            for j in range(len(first_list1)):
                second_list1 = first_list1[j].split(',')
                big_list1.extend(second_list1)
        for i2 in range(len(reference_list2)):
            first_list2 = reference_list2[i2].split(',')
            # print(first_list)
            for j2 in range(len(first_list2)):
                second_list2 = first_list2[j2].split(',')
                big_list2.extend(second_list2)

        # print(big_list1)
        print('第一个BOM共有', len(big_list1), '个位号')
        if len(set(big_list1)) != len(big_list1):
            print("BOM1中有重复值！")
            duplicates1 = list({x for x in big_list1 if big_list1.count(x) > 1})
            print('BOM1中重复位号：', duplicates1)
        else:
            print("BOM1没有重复值")
        print('\n')

        # print(big_list2)
        print('第二个BOM共有', len(big_list2), '个位号')
        if len(set(big_list2)) != len(big_list2):
            print("BOM2中有重复值！")
            duplicates2 = list({x for x in big_list2 if big_list2.count(x) > 1})
            print('BOM1中重复位号：', duplicates2)
        else:
            print("BOM2没有重复值")
        print('\n')

        diff1 = [x for x in big_list1 if x not in big_list2]
        diff2 = [x for x in big_list2 if x not in big_list1]
        if len(diff1) != 0:
            print('在BOM1但是不在BOM2的位号')
            print(diff1,'共',len(diff1),'个')
            df = pd.DataFrame(diff1, columns=["在BOM1但是不在BOM2的位号"])  # columns 是表头
            df.to_excel("在BOM1但是不在BOM2的位号.xlsx", index=False)
        else:
            print('BOM1中的位号都在BOM2中')
        if len(diff2) != 0:
            print('在BOM2但是不在BOM1的位号')
            print(diff2,'共',len(diff2),'个')
            df = pd.DataFrame(diff2, columns=["在BOM2但是不在BOM1的位号"])  # columns 是表头
            df.to_excel("在BOM2但是不在BOM1的位号.xlsx", index=False)
        else:
            print('BOM2中的位号都在BOM1中')