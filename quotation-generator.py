import requests
import math
import openpyxl
from openpyxl.drawing.image import Image
import pandas as pd
from pathlib import Path
from datetime import date
from decimal import Decimal, ROUND_HALF_UP, getcontext
import random
import string

getcontext().prec = 10  # 精度足夠即可

def excel_round(value, digits=-2):
    multiplier = Decimal('1e{}'.format(-digits))
    return int((Decimal(value) / multiplier).quantize(0, rounding=ROUND_HALF_UP) * multiplier)

mini價目表 = {
    'mini二節': { 90: 10900, 100:11200, 120: 11490, 150: 12400},
    'mini三節': { 90: 13400, 100:13700, 120: 13990, 150: 14900},
    'mini二節黑': { 90: 10900, 100:11200, 120: 11490, 150: 12400},
    'mini三節黑': { 90: 13400, 100:13700, 120: 13990, 150: 14900},
    'mini二節白': { 90: 10900, 100:11200, 120: 11490, 150: 12400},
    'mini三節白': { 90: 13400, 100:13700, 120: 13990, 150: 14900}
}

prime價目表 = {
    'prime二節': {
        100: {60: 12990, 80: None}, 
        120: {60: 12990, 80: 13500}, 
        150: {60: 13900, 80: 14500}, 
        180: {60: None, 80: 16100}
    },
    'prime三節': {
        100: {60: 15490, 80: None}, 
        120: {60: 15490, 80: 16000}, 
        150: {60: 16400, 80: 17000}, 
        180: {60: None, 80: 18600}
    },
    'prime二節黑': {
        100: {60: 12990, 80: None}, 
        120: {60: 12990, 80: 13500}, 
        150: {60: 13900, 80: 14500}, 
        180: {60: None, 80: 16100}
    },
    'prime三節黑': {
        100: {60: 15490, 80: None}, 
        120: {60: 15490, 80: 16000}, 
        150: {60: 16400, 80: 17000}, 
        180: {60: None, 80: 18600}
    },
    'prime二節白': {
        100: {60: 12990, 80: None}, 
        120: {60: 12990, 80: 13500}, 
        150: {60: 13900, 80: 14500}, 
        180: {60: None, 80: 16100}
    },
    'prime三節白': {
        100: {60: 15490, 80: None}, 
        120: {60: 15490, 80: 16000}, 
        150: {60: 16400, 80: 17000}, 
        180: {60: None, 80: 18600}
    },
    'prime二節灰': {
        100: {60: 12990, 80: None}, 
        120: {60: 12990, 80: 13500}, 
        150: {60: 13900, 80: 14500}, 
        180: {60: None, 80: 16100}
    },
    'prime三節灰': {
        100: {60: 15490, 80: None}, 
        120: {60: 15490, 80: 16000}, 
        150: {60: 16400, 80: 17000}, 
        180: {60: None, 80: 18600}
    }
}
force價目表 = {
    150: {80: 26000, 90: 26300},
    160: {80: 26500, 90: 26800},
    180: {80: 27600, 90: 27900},
    200: {80: 28300, 90: 28600},
    220: {80: 28800, 90: 29100}
}


客製價目表 = {
    (90,90):{
        (50,60):1990,
        (61,80):2500,
        (81,90):2800
    },
    (90.1,119.9):{
        (50,60):1990,
        (61,80):2500,
        (81,90):2800
    },
        (120,120):{
        (50,60):1990,
        (61,80):2500,
        (81,90):2800
    },
        (120.1,135):{
        (50,60):2400,
        (61,80):3000,
        (81,90):3300
    },
        (135.1,149.9):{
        (50,60):2900,
        (61,80):3500,
        (81,90):3800
    },
        (150,150):{
        (50,60):2900,
        (61,80):3500,
        (81,90):3800
    },
        (150.1,165):{
        (50,60):3400,
        (61,80):4000,
        (81,90):4300
    },
        (165.1,179.9):{
        (50,60):3900,
        (61,80):4500,
        (81,90):4800
    },
        (180,180):{
        (50,60):4400,
        (61,80):5100,
        (81,90):5400
    },
        (180.1,185):{
        (50,60):4600,
        (61,80):5500,
        (81,90):5800
    },
        (185.1,199.9):{
        (50,60):None,
        (61,80):6000,
        (81,90):6300
    },
        (200,200):{
        (50,60):None,
        (61,80):6000,
        (81,90):6300
    },
        (200.1,219.9):{
        (50,60):None,
        (61,80):6800,
        (81,90):7100
    },
        (220,220):{
        (50,60):None,
        (61,80):7000,
        (81,90):7300
    }
}

#規格品琥珀木價目表
規格琥珀木價目表={
    (50,70):{
        (90,120):9800,
        (120.1,150):12500,
        (150.1,180):14500,
        (180.1,210):16500,
        (210.1,240):18500
    },
    (71,90):{
        (120,149.9):11500,
        (150,179.9):14500,
        (180,209.9):17500,
        (210,239.9):20500,
        (240,240):22500
    }
}

                    
木種成本單價list={'栓木脂接':700,'栓木直拼':900,'白橡木脂接':850,'白橡木直拼':1330}
木種對客單價乘積list={'栓木脂接':1.5,'栓木直拼':1.5,'白橡木脂接':1.4,'白橡木直拼':1.4,'琥珀木':1.6}

桌腳_list={'prime三節':13500,'prime三節黑':13500,'prime三節白':13500,'prime三節灰':13500,
         'prime二節':11000,'prime二節黑':11000,'prime二節白':11000,'prime二節灰':11000,
         'mini三節':12000,'mini三節黑':12000,'mini三節白':12000,
         'mini二節':9500,'mini二節黑':9500,'mini二節白':9500,
         '固定桌腳':3980,'固定黑腳':3980,'固定白腳':3980,
         'force':23500,'force桌腳':23500,'force四柱桌腳':23500,'force四柱黑腳':23500,'force四柱白腳':23500,'':0}
顏色_list={'':0,'纖維板':0,'菸草橡木':0,'雪白柚木':0,'密西根楓木':0,'北歐白橡木':0,'典雅胡桃木':0,'歐風胡桃木':0,'黑':0,'白':0,'電競':0,'加拿大楓木':0, 
         '蜂巢板':1500,'黑雲岩':1500,'白雲岩':1500,'泥灰岩':1500,'安藤清水模':1500,'台灣柚木':1500,'北海道榆木':1500,'維吉尼亞楓木':1500,'安德森雪松':1500,'哥倫比亞胡桃':1500,}
形狀_list={'':0,'四方前上斜':0,'弧度上斜':0,'四方全平':0,'前凹':800,'後凹':800,'前凹+後凹':800,'小圓弧後凹':800,'前凹+小圓弧後凹':800,'四方前上斜+後凹':800,'四方前上斜+小圓弧後凹':800, '四角導圓':800}
製材所形狀_list={'':0,'四方全平':0,'四角導圓':0,'前自然邊':0,'後自然邊':0,'前後自然邊':0,'1號單邊上斜':0,'2號單邊上斜':0,'1號單邊上斜+四角導圓':0,'2號單邊上斜+四角導圓':0,'1號前後上斜':0,'2號前後上斜':0,'前凹':1500,'後凹':500,'前凹+後凹':2000}
         
while True:
    print('\n')
    try:
       報價 =input('請輸入規格或訂製或製材所或下單或報價單:')
    except Exception as e:
        print('輸入錯誤，原因:', e)
    else:

        if 報價 == '規格':
            try:
                桌寬 = float(input('請輸入桌寬:'))
                桌深 = float(input('請輸入桌深:'))
                桌腳 = input('請輸入桌腳:')
                顏色 = input('請輸入桌板顏色:')
                形狀 = input('請輸入桌板形狀:')

                顏色price=顏色_list[顏色]
                形狀price=形狀_list[形狀]  
                桌腳price=桌腳_list[桌腳]
            except Exception as e:
                print('輸入錯誤，原因:', e)
            else:       
                if 桌腳 == 'mini三節' or 桌腳 == 'mini二節' or 桌腳 == 'mini三節黑' or 桌腳 == 'mini三節白' or 桌腳 == 'mini二節黑' or 桌腳 == 'mini二節白':
                    try:
                        mini升降桌price = float(mini價目表[桌腳][桌寬])
                        顏色price=顏色_list[顏色]
                        形狀price=形狀_list[形狀] 
                        
                        if 桌深 != 60:
                            print('mini規格桌深須為60!!!')
                            continue
                    except Exception as e:
                        print('輸入錯誤，原因:', e) 
                    else:
                        total=mini升降桌price+顏色price+形狀price
        
                        print('和您報價')
                        print('%3.0f*%2.0f%s(%s)+%s=%5.0f'%(桌寬,桌深,顏色,形狀,桌腳,total))  
                    
                elif 桌腳 == 'prime三節' or 桌腳 == 'prime二節' or 桌腳 == 'prime三節黑' or 桌腳 == 'prime三節白' or 桌腳 == 'prime三節灰' or 桌腳 == 'prime二節黑' or 桌腳 == 'prime二節白' or 桌腳 =='prime二節灰':
                    try:       
                        prime升降桌price = float(prime價目表[桌腳][桌寬][桌深])
                        顏色price=顏色_list[顏色]
                        形狀price=形狀_list[形狀]
                    except Exception as e:
                        print('輸入錯誤，原因:', e)
                    else:
                        total=prime升降桌price+顏色price+形狀price
                        if 45<桌深<=60:
                            桌腳=桌腳+'(腳底座60公分)'
        
                        print('和您報價')
                        print('%3.0f*%2.0f%s(%s)+%s=%5.0f'%(桌寬,桌深,顏色,形狀,桌腳,total))
            
                elif 桌腳 == 'force' or 桌腳 =='force桌腳' or 桌腳 == 'force四柱桌腳' or 桌腳 == 'force四柱黑腳' or 桌腳 == 'force四柱白腳':
                    try:
                        force升降桌price = float(force價目表[桌寬][桌深])
                        顏色price=顏色_list[顏色]
                        形狀price=形狀_list[形狀]
                    except Exception as e:
                        print('輸入錯誤，原因:', e)
                    else:
                        total=force升降桌price+顏色price+形狀price
                
                        print('和您報價')
                        if 顏色price != 0:
                            print('訂%3.0f*%2.0f*4%s(%s)+%s=%5.0f'%(桌寬,桌深,顏色,形狀,桌腳,total))
                        else:
                            print('%3.0f*%2.0f%s(%s)+%s=%5.0f'%(桌寬,桌深,顏色,形狀,桌腳,total))
        
        elif 報價 == '訂製':
            try:
                桌寬 = float(input('請輸入桌寬:'))
                桌深 = math.ceil(float(input('請輸入桌深:')))
                桌腳=input('請輸入桌腳:')
                顏色=input('請輸入桌板顏色:')
                形狀=input('請輸入桌板形狀:')

                found = False
    
                for width_range, depth_dict in 客製價目表.items():
                    width_start, width_end = width_range
                    if width_start <= 桌寬 <= width_end:
        
                        for depth_range, value in depth_dict.items():
                            depth_start, depth_end = depth_range
                            if depth_start <= 桌深 <= depth_end:
                                price = value
                                found = True
                                break #找到寬度就停止查找
                        break #找到深度就停止查找
        
                if not found:
                    price = 0
                    print('查無價格!!!')
        
                桌腳price=桌腳_list[桌腳]
                顏色price=顏色_list[顏色]
                形狀price=形狀_list[形狀]
                
                def format_width_number(桌寬):
                    return str(int(桌寬)) if 桌寬 == int(桌寬) else str(桌寬)
                
            except Exception as e:
                print('輸入錯誤，原因:', e)
            else:
                total=price+桌腳price+顏色price+形狀price
                
                print('和您報價')
                if 桌腳price == 0 and 顏色price != 0 and 形狀price != 0:
                    total = total + 400
                    print(f'單購桌板訂{format_width_number(桌寬)}*{桌深}*4{顏色}({形狀})={total}')
                    print('\n''備註:')
                    print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                    print('2. 訂製約45天(含假日)')
                    print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')                     
                elif 桌腳price == 0 and 顏色price != 0:
                    total = total + 400
                    print(f'單購桌板訂{format_width_number(桌寬)}*{桌深}*4{顏色}({形狀})={total}')
                    print('請自行判斷是否為訂製!!')
                    print('\n''備註:')
                    print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                    print('2. 訂製約35天(含假日)')
                    print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主') 
                elif 桌腳price == 0:
                    total = total + 200
                    print(f'單購桌板訂{format_width_number(桌寬)}*{桌深}{顏色}({形狀})={total}')
                    print('請自行判斷是否為訂製!!')
                    print('\n''備註:')
                    print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                    print('2. 訂製約35天(含假日)')
                    print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主') 
                elif price == 0:
                    print('訂製尺寸有誤，注意桌深最小50')
                elif 桌寬<110 and (桌腳=='prime三節' or 桌腳=='prime二節' or 桌腳=='prime三節黑' or 桌腳=='prime二節黑'or 桌腳=='prime三節白' or 桌腳=='prime三節灰'or 桌腳=='prime二節灰' or 桌腳=='固定桌腳' or 桌腳 =='固定白腳' or 桌腳 == '固定黑腳') and 桌深<57.5:
                    桌腳=桌腳+'(短版+短側片+腳底座45公分)'
                    if 顏色price != 0 and 形狀price != 0:
                        if 形狀 == '四角導圓':
                            print(f'訂{format_width_number(桌寬)}*{桌深}*4{顏色}({形狀})+{桌腳}={total}')
                            print('\n''備註:')
                            print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                            print('2. 訂製約45天(含假日)')
                            print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主') 
                        else:
                            print('蜂巢板不可做凹!!')
                    elif 顏色price != 0:
                        print(f'訂{format_width_number(桌寬)}*{桌深}*4{顏色}({形狀})+{桌腳}={total}')
                        print('\n''備註:')
                        print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                        print('2. 訂製約35天(含假日)')
                        print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                    else:
                        if 形狀 == '四角導圓':
                            print('纖維板一定四角導圓，請刪除!!(否則報價有誤)')
                        else:
                            print(f'訂{format_width_number(桌寬)}*{桌深}{顏色}({形狀})+{桌腳}={total}')
                            print('\n''備註:')
                            print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                            print('2. 訂製約35天(含假日)')
                            print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
        
                elif 桌寬<110 and (桌腳=='prime三節' or 桌腳=='prime二節' or 桌腳=='prime三節黑' or 桌腳=='prime二節黑'or 桌腳=='prime三節白' or 桌腳=='prime二節白' or 桌腳=='prime三節灰'or 桌腳=='prime二節灰' or 桌腳=='固定桌腳' or 桌腳 =='固定白腳' or 桌腳 == '固定黑腳') and 桌深>=57.7:
                    if 57.5<=桌深<60:
                        桌腳=桌腳+'(短版+腳底座45公分)'
                        if 顏色price != 0 and 形狀price != 0:
                            if 形狀 == '四角導圓':
                                print(f'訂{format_width_number(桌寬)}*{桌深}*4{顏色}({形狀})+{桌腳}={total}')
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約45天(含假日)')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                            else:
                                print('蜂巢板不可做凹!!')
                        elif 顏色price != 0:
                            print(f'訂{format_width_number(桌寬)}*{桌深}*4{顏色}({形狀})+{桌腳}={total}')
                            print('\n''備註:')
                            print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                            print('2. 訂製約35天(含假日)')
                            print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                        else:
                            if 形狀 == '四角導圓':
                                print('纖維板一定四角導圓，請刪除!!(否則報價有誤)')
                            else:
                                print(f'訂{format_width_number(桌寬)}*{桌深}{顏色}({形狀})+{桌腳}={total}')
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約35天(含假日)')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                    elif 60<=桌深<68:
                        桌腳=桌腳+'(短版+腳底座60公分)'
                        if 顏色price != 0 and 形狀price != 0:
                            if 形狀 == '四角導圓':
                                print(f'訂{format_width_number(桌寬)}*{桌深}*4{顏色}({形狀})+{桌腳}={total}')
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約45天(含假日)')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                            else:
                                print('蜂巢板不可做凹!!')
                        elif 顏色price != 0:
                            print(f'訂{format_width_number(桌寬)}*{桌深}*4{顏色}({形狀})+{桌腳}={total}')
                            print('\n''備註:')
                            print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                            print('2. 訂製約35天(含假日)')
                            print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                        else:
                            if 形狀 == '四角導圓':
                                print('纖維板一定四角導圓，請刪除!!(否則報價有誤)')
                            else:
                                print(f'訂{format_width_number(桌寬)}*{桌深}{顏色}({形狀})+{桌腳}={total}')
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約35天(含假日)')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                    else:
                        桌腳=桌腳+'(短版)'
                        if 顏色price != 0 and 形狀price != 0:
                            if 形狀 == '四角導圓':
                                print(f'訂{format_width_number(桌寬)}*{桌深}*4{顏色}({形狀})+{桌腳}={total}')
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約45天(含假日)')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                            else:
                                print('蜂巢板不可做凹!!')
                        elif 顏色price != 0:
                            print(f'訂{format_width_number(桌寬)}*{桌深}*4{顏色}({形狀})+{桌腳}={total}')
                            print('\n''備註:')
                            print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                            print('2. 訂製約35天(含假日)')
                            print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                        else:
                            if 形狀 == '四角導圓':
                                print('纖維板一定四角導圓，請刪除!!(否則報價有誤)')
                            else:
                                print(f'訂{format_width_number(桌寬)}*{桌深}{顏色}({形狀})+{桌腳}={total}')
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約35天(含假日)')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
               
                elif 桌深<68 and (桌腳=='prime三節' or 桌腳=='prime二節' or 桌腳=='prime三節黑' or 桌腳=='prime二節黑'or 桌腳=='prime三節白' or 桌腳=='prime二節白' or 桌腳=='prime三節灰'or 桌腳=='prime二節灰'or 桌腳=='固定桌腳' or 桌腳 =='固定白腳' or 桌腳 == '固定黑腳'):
                    if 57.5<=桌深<68:
                        桌腳=桌腳+'(腳底座60公分)'
                        if 顏色price != 0 and 形狀price != 0:
                            if 形狀 == '四角導圓':
                                print(f'訂{format_width_number(桌寬)}*{桌深}*4{顏色}({形狀})+{桌腳}={total}')
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約45天(含假日)')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                            else:
                                print('蜂巢板不可做凹!!')
                        elif 顏色price != 0:
                            print(f'訂{format_width_number(桌寬)}*{桌深}*4{顏色}({形狀})+{桌腳}={total}')
                            print('\n''備註:')
                            print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                            print('2. 訂製約35天(含假日)')
                            print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                        else:
                            if 形狀 == '四角導圓':
                                print('纖維板一定四角導圓，請刪除!!(否則報價有誤)')
                            else:
                                print(f'訂{format_width_number(桌寬)}*{桌深}{顏色}({形狀})+{桌腳}={total}')
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約35天(含假日)')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                    elif 桌深<57.5:
                        桌腳=桌腳+'(短側片+腳底座45公分)'
                        if 顏色price != 0 and 形狀price != 0:
                            if 形狀 == '四角導圓':
                                print(f'訂{format_width_number(桌寬)}*{桌深}*4{顏色}({形狀})+{桌腳}={total}')
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約45天(含假日)')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                            else:
                                print('蜂巢板不可做凹!!')
                        elif 顏色price != 0:
                            print(f'訂{format_width_number(桌寬)}*{桌深}*4{顏色}({形狀})+{桌腳}={total}')
                            print('\n''備註:')
                            print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                            print('2. 訂製約35天(含假日)')
                            print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                        else:
                            if 形狀 == '四角導圓':
                                print('纖維板一定四角導圓，請刪除!!(否則報價有誤)')
                            else:
                                print(f'訂{format_width_number(桌寬)}*{桌深}{顏色}({形狀})+{桌腳}={total}')
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約35天(含假日)')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                                
                elif 桌深<60 and (桌腳=='mini三節' or 桌腳=='mini二節' or 桌腳=='mini三節白' or 桌腳=='mini二節白' or 桌腳=='mini三節黑' or 桌腳=='mini二節黑'or 桌腳=='固定桌腳' or 桌腳 =='固定白腳' or 桌腳 == '固定黑腳'):
                    桌腳 = 桌腳+'(腳底座45公分)'
                    print(f'訂{format_width_number(桌寬)}*{桌深}{顏色}({形狀})+{桌腳}={total}')
                    print('\n''備註:')
                    print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                    print('2. 訂製約35天(含假日)')
                    print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                    
                elif 桌深<72 and 桌腳=='force':
                    print('桌深無法安裝force!')
        
                else:
                    if 顏色price != 0 and 形狀price != 0:
                        if 形狀 == '四角導圓':
                            print(f'訂{format_width_number(桌寬)}*{桌深}*4{顏色}({形狀})+{桌腳}={total}')
                            print('\n''備註:')
                            print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                            print('2. 訂製約45天(含假日)')
                            print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                        else:
                            print('蜂巢板不可做凹!!')
                    elif 顏色price != 0:
                        print(f'訂{format_width_number(桌寬)}*{桌深}*4{顏色}({形狀})+{桌腳}={total}')
                        print('\n''備註:')
                        print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                        print('2. 訂製約35天(含假日)')
                        print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                    else:
                        if 形狀 == '四角導圓':
                            print('纖維板一定四角導圓，請刪除!!(否則報價有誤)')
                        else:
                            print(f'訂{format_width_number(桌寬)}*{桌深}{顏色}({形狀})+{桌腳}={total}')
                            print('\n''備註:')
                            print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                            print('2. 訂製約35天(含假日)')
                            print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
            
        elif 報價 == '製材所':
            try:
                木種=input('請輸入木種:')
            except Exception as e:
                print('輸入錯誤，原因:', e)
            else:
                if 木種 =='栓木脂接' or 木種 =='栓木直拼':
                    try:
                        桌寬=float(input('請輸入桌寬:'))
                        桌深=float(input('請輸入桌深:'))
                        厚度=float(input('請輸入厚度(2.7/3.5/4.5):'))
                        桌腳=input('請輸入桌腳:') 
                        製材所形狀=input('請輸入桌板形狀:')
                        桌腳price=桌腳_list[桌腳] 
                        製材所形狀price=製材所形狀_list[製材所形狀]

                        # ✅ 加入防呆：若輸入厚度為 3.3，自動修正為 3.5
                        if abs(厚度 - 3.3) < 0.05:
                            厚度 = 3.5
                    except Exception as e:
                        print('輸入錯誤，原因:', e)
                    else:
                        if 桌腳 == '':
                            # 這一行替換為 Decimal 計算（避免 float 誤差）
                            基礎價格 = Decimal(str(桌寬)) * Decimal(str(桌深)) * Decimal(str(厚度)) / Decimal('2700') * Decimal(str(木種成本單價list[木種])) * Decimal(str(木種對客單價乘積list[木種]))
                            桌板price = excel_round(基礎價格, -2)
                            total=桌板price+製材所形狀price+400

                            if 製材所形狀 =='':
                                製材所形狀 = '四方全平'
                            print('和您報價')
                            print('(製材所)-單購桌板訂%3.0f*%3.0f*%1.1f%s(%s)=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,total))
                            print('\n''備註:')
                            print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                            print('2. 訂製約45工作天')
                            print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                            
                        elif 桌深 >= 81 and 厚度<=2.7:
                                print('無法製作!!桌深厚度81cm以上，厚度需為3.3以上')
    
                        elif 桌寬<110 and (桌腳=='prime三節' or 桌腳=='prime二節' or 桌腳=='prime三節黑' or 桌腳=='prime二節黑'or 桌腳=='prime三節白' or 桌腳=='prime三節灰'or 桌腳=='prime二節灰' or 桌腳=='固定桌腳' or 桌腳 =='固定白腳' or 桌腳 == '固定黑腳') and 桌深<57.5:
                            桌腳=桌腳+'(短版+短側片+腳底座45公分)'
                            基礎價格 = Decimal(str(桌寬)) * Decimal(str(桌深)) * Decimal(str(厚度)) / Decimal('2700') * Decimal(str(木種成本單價list[木種])) * Decimal(str(木種對客單價乘積list[木種]))
                            桌板price = excel_round(基礎價格, -2)
                            total=桌板price+製材所形狀price+桌腳price

                            if 製材所形狀 =='':
                                製材所形狀 = '四方全平'
                            print('和您報價')
                            print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                            print('\n''備註:')
                            print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                            print('2. 訂製約45工作天')
                            print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
        
                        elif 桌寬<110 and (桌腳=='prime三節' or 桌腳=='prime二節' or 桌腳=='prime三節黑' or 桌腳=='prime二節黑'or 桌腳=='prime三節白' or 桌腳=='prime二節白' or 桌腳=='prime三節灰'or 桌腳=='prime二節灰' or 桌腳=='固定桌腳' or 桌腳 =='固定白腳' or 桌腳 == '固定黑腳') and 桌深>=57.7:
                            if 57.5<=桌深<60:
                                桌腳=桌腳+'(短版+腳底座45公分)'
                                基礎價格 = Decimal(str(桌寬)) * Decimal(str(桌深)) * Decimal(str(厚度)) / Decimal('2700') * Decimal(str(木種成本單價list[木種])) * Decimal(str(木種對客單價乘積list[木種]))
                                桌板price = excel_round(基礎價格, -2)
                                total=桌板price+製材所形狀price+桌腳price
    
                                if 製材所形狀 =='':
                                    製材所形狀 = '四方全平'
                                print('和您報價')
                                print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約45工作天')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                                    
                            elif 60<=桌深<68:
                                桌腳=桌腳+'(短版+腳底座60公分)'
                                基礎價格 = Decimal(str(桌寬)) * Decimal(str(桌深)) * Decimal(str(厚度)) / Decimal('2700') * Decimal(str(木種成本單價list[木種])) * Decimal(str(木種對客單價乘積list[木種]))
                                桌板price = excel_round(基礎價格, -2)
                                total=桌板price+製材所形狀price+桌腳price
    
                                if 製材所形狀 =='':
                                    製材所形狀 = '四方全平'
                                print('和您報價')
                                print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約45工作天')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                            else:
                                桌腳=桌腳+'(短版)'
                                基礎價格 = Decimal(str(桌寬)) * Decimal(str(桌深)) * Decimal(str(厚度)) / Decimal('2700') * Decimal(str(木種成本單價list[木種])) * Decimal(str(木種對客單價乘積list[木種]))
                                桌板price = excel_round(基礎價格, -2)
                                total=桌板price+製材所形狀price+桌腳price
    
                                if 製材所形狀 =='':
                                    製材所形狀 = '四方全平'
                                print('和您報價')
                                print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約45工作天')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                        
               
                        elif 桌深<68 and (桌腳=='prime三節' or 桌腳=='prime二節' or 桌腳=='prime三節黑' or 桌腳=='prime二節黑'or 桌腳=='prime三節白' or 桌腳=='prime二節白' or 桌腳=='prime三節灰'or 桌腳=='prime二節灰'or 桌腳=='固定桌腳' or 桌腳 =='固定白腳' or 桌腳 == '固定黑腳'):
                            if 57.5<=桌深<68:
                                桌腳=桌腳+'(腳底座60公分)'
                                基礎價格 = Decimal(str(桌寬)) * Decimal(str(桌深)) * Decimal(str(厚度)) / Decimal('2700') * Decimal(str(木種成本單價list[木種])) * Decimal(str(木種對客單價乘積list[木種]))
                                桌板price = excel_round(基礎價格, -2)
                                total=桌板price+製材所形狀price+桌腳price
        
                                if 製材所形狀 =='':
                                    製材所形狀 = '四方全平'
                                print('和您報價')
                                print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約45工作天')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                            elif 桌深<57.5:
                                桌腳=桌腳+'(短側片+腳底座45公分)'
                                基礎價格 = Decimal(str(桌寬)) * Decimal(str(桌深)) * Decimal(str(厚度)) / Decimal('2700') * Decimal(str(木種成本單價list[木種])) * Decimal(str(木種對客單價乘積list[木種]))
                                桌板price = excel_round(基礎價格, -2)
                                total=桌板price+製材所形狀price+桌腳price
        
                                if 製材所形狀 =='':
                                    製材所形狀 = '四方全平'
                                print('和您報價')
                                print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約45工作天')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                                
                        elif 桌深<60 and (桌腳=='mini三節' or 桌腳=='mini二節' or 桌腳=='mini三節白' or 桌腳=='mini二節白' or 桌腳=='mini三節黑' or 桌腳=='mini二節黑'or 桌腳=='固定桌腳' or 桌腳 =='固定白腳' or 桌腳 == '固定黑腳'):
                            桌腳 = 桌腳+'(腳底座45公分)'
                            基礎價格 = Decimal(str(桌寬)) * Decimal(str(桌深)) * Decimal(str(厚度)) / Decimal('2700') * Decimal(str(木種成本單價list[木種])) * Decimal(str(木種對客單價乘積list[木種]))
                            桌板price = excel_round(基礎價格, -2)
                            total=桌板price+製材所形狀price+桌腳price
        
                            if 製材所形狀 =='':
                                製材所形狀 = '四方全平'
                            print('和您報價')
                            print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                            print('\n''備註:')
                            print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                            print('2. 訂製約45工作天')
                            print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                            
                        elif 桌深<72 and 桌腳=='force':
                            print('桌深無法安裝force!')
                        
                        else:
                            基礎價格 = Decimal(str(桌寬)) * Decimal(str(桌深)) * Decimal(str(厚度)) / Decimal('2700') * Decimal(str(木種成本單價list[木種])) * Decimal(str(木種對客單價乘積list[木種]))
                            桌板price = excel_round(基礎價格, -2)
                            total=桌板price+製材所形狀price+桌腳price

                            if 製材所形狀 =='':
                                製材所形狀 = '四方全平'
                            print('和您報價')
                            print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                            print('\n''備註:')
                            print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                            print('2. 訂製約45工作天')
                            print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                            
                elif 木種 =='白橡木脂接' or 木種 =='白橡木直拼':
                    try:
                        桌寬=float(input('請輸入桌寬:'))
                        桌深=float(input('請輸入桌深:'))
                        厚度=float(input('請輸入厚度(2.7/3.3/4.5):'))
                        桌腳=input('請輸入桌腳:') 
                        製材所形狀=input('請輸入桌板形狀:')
                        桌腳price=桌腳_list[桌腳] 
                        製材所形狀price=製材所形狀_list[製材所形狀]

                        #✅ 加入防呆：若輸入厚度為 3.3，自動修正為 3.5
                        if abs(厚度 - 3.3) < 0.05:
                            厚度 = 3.5
                    except Exception as e:
                        print('輸入錯誤，原因:', e)
                    else:
                        if 桌腳 == '':
                            # 這一行替換為 Decimal 計算（避免 float 誤差）
                            基礎價格 = Decimal(str(桌寬)) * Decimal(str(桌深)) * Decimal(str(厚度)) / Decimal('2700') * Decimal(str(木種成本單價list[木種])) * Decimal(str(木種對客單價乘積list[木種]))
                            桌板price = excel_round(基礎價格, -2)
                            total=桌板price+製材所形狀price+400

                            if 製材所形狀 =='':
                                製材所形狀 = '四方全平'
                            print('和您報價')
                            print('(製材所)-單購桌板訂%3.0f*%3.0f*%1.1f%s(%s)=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,total))
                            print('\n''備註:')
                            print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                            print('2. 訂製約45工作天')
                            print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                            
                        elif 桌深 >= 81 and 厚度<=2.7:
                                print('無法製作!!桌深厚度81cm以上，厚度需為3.3以上')
    
                        elif 桌寬<110 and (桌腳=='prime三節' or 桌腳=='prime二節' or 桌腳=='prime三節黑' or 桌腳=='prime二節黑'or 桌腳=='prime三節白' or 桌腳=='prime三節灰'or 桌腳=='prime二節灰' or 桌腳=='固定桌腳' or 桌腳 =='固定白腳' or 桌腳 == '固定黑腳') and 桌深<57.5:
                            桌腳=桌腳+'(短版+短側片+腳底座45公分)'
                            基礎價格 = Decimal(str(桌寬)) * Decimal(str(桌深)) * Decimal(str(厚度)) / Decimal('2700') * Decimal(str(木種成本單價list[木種])) * Decimal(str(木種對客單價乘積list[木種]))
                            桌板price = excel_round(基礎價格, -2)
                            total=桌板price+製材所形狀price+桌腳price

                            if 製材所形狀 =='':
                                製材所形狀 = '四方全平'
                            print('和您報價')
                            print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                            print('\n''備註:')
                            print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                            print('2. 訂製約45工作天')
                            print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
        
                        elif 桌寬<110 and (桌腳=='prime三節' or 桌腳=='prime二節' or 桌腳=='prime三節黑' or 桌腳=='prime二節黑'or 桌腳=='prime三節白' or 桌腳=='prime二節白' or 桌腳=='prime三節灰'or 桌腳=='prime二節灰' or 桌腳=='固定桌腳' or 桌腳 =='固定白腳' or 桌腳 == '固定黑腳') and 桌深>=57.7:
                            if 57.5<=桌深<60:
                                桌腳=桌腳+'(短版+腳底座45公分)'
                                基礎價格 = Decimal(str(桌寬)) * Decimal(str(桌深)) * Decimal(str(厚度)) / Decimal('2700') * Decimal(str(木種成本單價list[木種])) * Decimal(str(木種對客單價乘積list[木種]))
                                桌板price = excel_round(基礎價格, -2)
                                total=桌板price+製材所形狀price+桌腳price
    
                                if 製材所形狀 =='':
                                    製材所形狀 = '四方全平'
                                print('和您報價')
                                print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約45工作天')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                                    
                            elif 60<=桌深<68:
                                桌腳=桌腳+'(短版+腳底座60公分)'
                                基礎價格 = Decimal(str(桌寬)) * Decimal(str(桌深)) * Decimal(str(厚度)) / Decimal('2700') * Decimal(str(木種成本單價list[木種])) * Decimal(str(木種對客單價乘積list[木種]))
                                桌板price = excel_round(基礎價格, -2)
                                total=桌板price+製材所形狀price+桌腳price
    
                                if 製材所形狀 =='':
                                    製材所形狀 = '四方全平'
                                print('和您報價')
                                print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約45工作天')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                            else:
                                桌腳=桌腳+'(短版)'
                                基礎價格 = Decimal(str(桌寬)) * Decimal(str(桌深)) * Decimal(str(厚度)) / Decimal('2700') * Decimal(str(木種成本單價list[木種])) * Decimal(str(木種對客單價乘積list[木種]))
                                桌板price = excel_round(基礎價格, -2)
                                total=桌板price+製材所形狀price+桌腳price
    
                                if 製材所形狀 =='':
                                    製材所形狀 = '四方全平'
                                print('和您報價')
                                print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約45工作天')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                        
               
                        elif 桌深<68 and (桌腳=='prime三節' or 桌腳=='prime二節' or 桌腳=='prime三節黑' or 桌腳=='prime二節黑'or 桌腳=='prime三節白' or 桌腳=='prime二節白' or 桌腳=='prime三節灰'or 桌腳=='prime二節灰'or 桌腳=='固定桌腳' or 桌腳 =='固定白腳' or 桌腳 == '固定黑腳'):
                            if 57.5<=桌深<68:
                                桌腳=桌腳+'(腳底座60公分)'
                                基礎價格 = Decimal(str(桌寬)) * Decimal(str(桌深)) * Decimal(str(厚度)) / Decimal('2700') * Decimal(str(木種成本單價list[木種])) * Decimal(str(木種對客單價乘積list[木種]))
                                桌板price = excel_round(基礎價格, -2)
                                total=桌板price+製材所形狀price+桌腳price
        
                                if 製材所形狀 =='':
                                    製材所形狀 = '四方全平'
                                print('和您報價')
                                print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約45工作天')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                            elif 桌深<57.5:
                                桌腳=桌腳+'(短側片+腳底座45公分)'
                                基礎價格 = Decimal(str(桌寬)) * Decimal(str(桌深)) * Decimal(str(厚度)) / Decimal('2700') * Decimal(str(木種成本單價list[木種])) * Decimal(str(木種對客單價乘積list[木種]))
                                桌板price = excel_round(基礎價格, -2)
                                total=桌板price+製材所形狀price+桌腳price
        
                                if 製材所形狀 =='':
                                    製材所形狀 = '四方全平'
                                print('和您報價')
                                print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約45工作天')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                                
                        elif 桌深<60 and (桌腳=='mini三節' or 桌腳=='mini二節' or 桌腳=='mini三節白' or 桌腳=='mini二節白' or 桌腳=='mini三節黑' or 桌腳=='mini二節黑'or 桌腳=='固定桌腳' or 桌腳 =='固定白腳' or 桌腳 == '固定黑腳'):
                            桌腳 = 桌腳+'(腳底座45公分)'
                            基礎價格 = Decimal(str(桌寬)) * Decimal(str(桌深)) * Decimal(str(厚度)) / Decimal('2700') * Decimal(str(木種成本單價list[木種])) * Decimal(str(木種對客單價乘積list[木種]))
                            桌板price = excel_round(基礎價格, -2)
                            total=桌板price+製材所形狀price+桌腳price
        
                            if 製材所形狀 =='':
                                製材所形狀 = '四方全平'
                            print('和您報價')
                            print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                            print('\n''備註:')
                            print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                            print('2. 訂製約45工作天')
                            print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                            
                        elif 桌深<72 and 桌腳=='force':
                            print('桌深無法安裝force!')
                        
                        else:
                            基礎價格 = Decimal(str(桌寬)) * Decimal(str(桌深)) * Decimal(str(厚度)) / Decimal('2700') * Decimal(str(木種成本單價list[木種])) * Decimal(str(木種對客單價乘積list[木種]))
                            桌板price = excel_round(基礎價格, -2)
                            total=桌板price+製材所形狀price+桌腳price

                            if 製材所形狀 =='':
                                製材所形狀 = '四方全平'
                            print('和您報價')
                            print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                            print('\n''備註:')
                            print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                            print('2. 訂製約45工作天')
                            print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                            
                elif 木種 == '琥珀木':
                    try:
                        桌寬=float(input('請輸入桌寬:'))
                        桌深=float(input('請輸入桌深:'))
                        厚度=float(input('厚度請輸入4.5:'))
                        桌腳=input('請輸入桌腳:') 
                        製材所形狀=input('請輸入桌板形狀:')
             
                        for wood_depth_range, wood_width_dict in 規格琥珀木價目表.items():
                            wood_depth_start, wood_depth_end = wood_depth_range
                            if wood_depth_start <= 桌深 <= wood_depth_end:
        
                                for wood_width_range, wood_value in wood_width_dict.items():
                                    wood_width_start, wood_width_end = wood_width_range
                                    if wood_width_start <= 桌寬 <= wood_width_end:
                                        price = wood_value
                                        break #找到寬度就停止查找
                                break #找到深度就停止查找
        
                        else:
                            print('無價格')

                        桌板price = price*木種對客單價乘積list[木種]
                        製材所形狀price=製材所形狀_list[製材所形狀]
                    except Exception as e:
                        print('輸入錯誤，原因:', e)
                    else:
                        if 桌腳 == '':
                            total=桌板price+製材所形狀price+400

                            if 製材所形狀 =='':
                                製材所形狀 = '四方全平'
                            print('和您報價')
                            print('(製材所)-單購桌板訂%3.0f*%3.0f*%1.1f%s(%s)=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,total))
                            print('\n''備註:')
                            print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                            print('2. 訂製約45工作天')
                            print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                            
                        elif 桌深 >= 81 and 厚度<=2.7:
                                print('無法製作!!桌深厚度81cm以上，厚度需為3.3以上')
    
                        elif 桌寬<110 and (桌腳=='prime三節' or 桌腳=='prime二節' or 桌腳=='prime三節黑' or 桌腳=='prime二節黑'or 桌腳=='prime三節白' or 桌腳=='prime三節灰'or 桌腳=='prime二節灰' or 桌腳=='固定桌腳' or 桌腳 =='固定白腳' or 桌腳 == '固定黑腳') and 桌深<57.5:
                            桌腳=桌腳+'(短版+短側片+腳底座45公分)'
                            total=桌板price+製材所形狀price+桌腳price

                            if 製材所形狀 =='':
                                製材所形狀 = '四方全平'
                            print('和您報價')
                            print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                            print('\n''備註:')
                            print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                            print('2. 訂製約45工作天')
                            print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
        
                        elif 桌寬<110 and (桌腳=='prime三節' or 桌腳=='prime二節' or 桌腳=='prime三節黑' or 桌腳=='prime二節黑'or 桌腳=='prime三節白' or 桌腳=='prime二節白' or 桌腳=='prime三節灰'or 桌腳=='prime二節灰' or 桌腳=='固定桌腳' or 桌腳 =='固定白腳' or 桌腳 == '固定黑腳') and 桌深>=57.7:
                            if 57.5<=桌深<60:
                                桌腳=桌腳+'(短版+腳底座45公分)'
                                total=桌板price+製材所形狀price+桌腳price
    
                                if 製材所形狀 =='':
                                    製材所形狀 = '四方全平'
                                print('和您報價')
                                print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約45工作天')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                                    
                            elif 60<=桌深<68:
                                桌腳=桌腳+'(短版+腳底座60公分)'
                                total=桌板price+製材所形狀price+桌腳price
    
                                if 製材所形狀 =='':
                                    製材所形狀 = '四方全平'
                                print('和您報價')
                                print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約45工作天')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                            else:
                                桌腳=桌腳+'(短版)'
                                total=桌板price+製材所形狀price+桌腳price
    
                                if 製材所形狀 =='':
                                    製材所形狀 = '四方全平'
                                print('和您報價')
                                print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約45工作天')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                        
               
                        elif 桌深<68 and (桌腳=='prime三節' or 桌腳=='prime二節' or 桌腳=='prime三節黑' or 桌腳=='prime二節黑'or 桌腳=='prime三節白' or 桌腳=='prime二節白' or 桌腳=='prime三節灰'or 桌腳=='prime二節灰'or 桌腳=='固定桌腳' or 桌腳 =='固定白腳' or 桌腳 == '固定黑腳'):
                            if 57.5<=桌深<68:
                                桌腳=桌腳+'(腳底座60公分)'
                                total=桌板price+製材所形狀price+桌腳price
        
                                if 製材所形狀 =='':
                                    製材所形狀 = '四方全平'
                                print('和您報價')
                                print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約45工作天')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                            elif 桌深<57.5:
                                桌腳=桌腳+'(短側片+腳底座45公分)'
                                total=桌板price+製材所形狀price+桌腳price
        
                                if 製材所形狀 =='':
                                    製材所形狀 = '四方全平'
                                print('和您報價')
                                print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                                print('\n''備註:')
                                print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                                print('2. 訂製約45工作天')
                                print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                                
                        elif 桌深<60 and (桌腳=='mini三節' or 桌腳=='mini二節' or 桌腳=='mini三節白' or 桌腳=='mini二節白' or 桌腳=='mini三節黑' or 桌腳=='mini二節黑'or 桌腳=='固定桌腳' or 桌腳 =='固定白腳' or 桌腳 == '固定黑腳'):
                            桌腳 = 桌腳+'(腳底座45公分)'
                            total=桌板price+製材所形狀price+桌腳price
        
                            if 製材所形狀 =='':
                                製材所形狀 = '四方全平'
                            print('和您報價')
                            print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                            print('\n''備註:')
                            print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                            print('2. 訂製約45工作天')
                            print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                            
                        elif 桌深<72 and 桌腳=='force':
                            print('桌深無法安裝force!')
                        
                        else:
                            total=桌板price+製材所形狀price+桌腳price

                            if 製材所形狀 =='':
                                製材所形狀 = '四方全平'
                            print('和您報價')
                            print('(製材所)-訂%3.0f*%3.0f*%1.1f%s(%s)+%s=%5.0f' % (桌寬,桌深,厚度,木種,製材所形狀,桌腳,total))
                            print('\n''備註:')
                            print('1. 以上金額含運(不含宜花東地區)、不含安裝')
                            print('2. 訂製約45工作天')
                            print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                else:
                    print('木種輸入錯誤!!')

        elif 報價 == '報價單':
            # === 0. 路徑設定 ===============================================
            base_dir = Path.cwd()
            TEMPLATE   = base_dir / "FUNTE電動升降桌報價單-空白報價單.xlsx"
            PRODUCT_DB = base_dir / "報價單產品.xlsx"
            
            # === 1. 讀產品資料表 ==========================================
            items_df = pd.read_excel(PRODUCT_DB, header=14)
            items_df.columns = items_df.columns.str.strip()
            items_df.set_index("Item No.", inplace=True)
            
            # === 2. 客戶資料 ==============================================
            customer = {
                "name": input("姓名／公司名／統編：").strip(),
                "tel":  input("電話：").strip(),
                "addr": input("地址：").strip()
            }
            
            # === 3. 多品項輸入 ============================================
            orders = []
            while True:
                prod = input("產品名稱（Enter 結束）：").strip()
                if not prod:
                    break
                if prod not in items_df.index:
                    print("⚠️ 找不到此產品，請重輸")
                    continue
                orders.append({"row": items_df.loc[prod]})
            
            if not orders:
                print("⚠️ 沒有輸入任何產品，程式結束")
                exit()
            
            # === 4. 讀範本，寫客戶資料 ===================================
            wb = openpyxl.load_workbook(TEMPLATE)
            ws = wb.active
            ws["A9"].value  = f"Messrs：{customer['name']}"
            ws["A10"].value = f"TEL：{customer['tel']}"
            ws["A11"].value = f"ADDR：{customer['addr']}"
            ws["C11"].value = f"DATE：{date.today():%Y/%m/%d}"
            
            # === 5. 插入空白列：把原本總計/備註往下推 =====================
            HEADER_ROW  = 15        # 表頭在第 15 列
            start_row   = HEADER_ROW + 1            # 第一筆資料 → 16
            extra_rows  = len(orders) - 1
            if extra_rows > 0:
                ws.insert_rows(
                    idx=start_row + 1, amount=extra_rows)
            
            # === 6. 寫產品列 & 圖片 =====================================
            for i, od in enumerate(orders):
                row = start_row + i
                r = od["row"]
            
                #ws.cell(row, 1, r["Item No."])
                ws.cell(row, 2, r["Description"])
                ws.cell(row, 3).value = f"=ROUND({r["Price (NTD)"]}, 0)" 
                ws.cell(row, 5, f"=C{row}*D{row}")             # E 欄小計
            
                # 插圖（路徑用 base_dir）
                        
                item_no = str(r.name).strip()        # 取得「大洞洞板」
                img_path = base_dir / "報價單產品圖檔" / f"{item_no}.JPG"
                if img_path.exists():
                    img = Image(str(img_path))
                    img.width, img.height = 120, 90
                    ws.column_dimensions["A"].width = 18  # 大約適合120px圖片
                    ws.add_image(img, f"A{row}")
                else:
                    print(f"⚠️ 找不到圖片：{img_path}")
            
            # === 7. 重寫總計列公式 =======================================
            total_row = start_row + len(orders)
            ws.cell(total_row, 1, "總計新台幣（含稅）")
            ws.cell(total_row, 5, f"=SUM(E{start_row}:E{total_row-1})")
            
            # === 8. 另存檔案 =============================================
            safe_name = customer["name"].replace(" ", "_")
            outfile = base_dir / f"FUNTE電動升降桌報價單_{safe_name}_{date.today():%Y%m%d}.xlsx"
            wb.save(outfile)
            
            print(f"\n✅ 報價單已產生：{outfile}")
            
        elif 報價 == '下單':            
            ACCESS_TOKEN = 'SHOPLINE_API_KEY'
            API_BASE_URL = 'https://open.shopline.io/v1'

            headers = {
                'Authorization': f'Bearer {ACCESS_TOKEN}',
                'Content-Type': 'application/json'
            }

            try:
                custom_name = input('請輸入姓名: ')
                custom_product = input('請輸入商品名稱: ')
                custom_price = int(input('請輸入價格: '))
                custom_sku = input('請輸入物料(CD01/CD02):')
                #custom_url_code = input('請輸入連結編碼: ')
            except Exception as e:
                print('輸入錯誤，原因:', e)
            else:
                # 自動產生亂碼連結（14碼）
                custom_url_code = ''.join(random.choices(string.ascii_lowercase + string.digits, k=14))
                payload = {
                    "product": {
                        "title_translations": {
                            "zh-hant": f"【{custom_name}】-{custom_product}"
                        },
                        "category_ids":[
                            "5f3e6edecbccfd003c3fc43f","6620f0d9de0f7400174633cd"
                        ],
                        "summary_translations": {
                            "zh-hant": f"🔺製作天數不含配送時間，商品完成後將依訂單順序安排出貨"
                        },
                        "price": 99999,  # 原價，可視情況調整
                        "price_sale": custom_price,
                        "unlimited_quantity": True,
                        "sku": custom_sku,
                        "weight": 50,
                        "status": "hidden",
                        "blacklisted_payment_ids": [
                            "59d302ce080f065a31000066"
                        ],
                        "images": [
                            "SHOPLINE_IMAGES"
                        ],
                        "link": custom_url_code
                    }
                }

                product_url = f'{API_BASE_URL}/products'

                response = requests.post(product_url, headers=headers, json=payload)

                if response.status_code == 201:
                    base_url = 'https://www.funtetw.com/products/'
                    print(f'✅ 已生成連結：{base_url}{custom_url_code}')

                else:
                    print(f'❌ 建立失敗，狀態碼：{response.status_code}')
                    print(response.text)

        else:
            print('選項輸入錯誤!!')
    
input('請按enter鍵結束')
    #again = input("要重新報價嗎？(y/n)：")  
    #if again.lower() != 'y':
            #break