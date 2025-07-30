import requests
import math

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
        100: {60: 12990, 80: None}, 
        120: {60: 12990, 80: 13500}, 
        150: {60: 13900, 80: 14500}, 
        180: {60: None, 80: 16100}
    },
    'prime二節黑': {
        100: {60: 12990, 80: None}, 
        120: {60: 12990, 80: 13500}, 
        150: {60: 13900, 80: 14500}, 
        180: {60: None, 80: 16100}
    },
    'prime三節黑': {
        100: {60: 12990, 80: None}, 
        120: {60: 12990, 80: 13500}, 
        150: {60: 13900, 80: 14500}, 
        180: {60: None, 80: 16100}
    },
    'prime二節白': {
        100: {60: 12990, 80: None}, 
        120: {60: 12990, 80: 13500}, 
        150: {60: 13900, 80: 14500}, 
        180: {60: None, 80: 16100}
    },
    'prime三節白': {
        100: {60: 12990, 80: None}, 
        120: {60: 12990, 80: 13500}, 
        150: {60: 13900, 80: 14500}, 
        180: {60: None, 80: 16100}
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
        (121.1,135):{
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
        (120,149.9):9800,
        (150,179.9):12500,
        (180,209.9):14500,
        (210,239.9):16500,
        (240,240):18500
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

桌腳_list={'prime三節':13500,'prime三節黑':13500,'prime三節白':13500,
         'prime二節':11000,'prime二節黑':11000,'prime二節白':11000,
         'mini三節':12000,'mini三節黑':12000,'mini三節白':12000,
         'mini二節':9500,'mini二節黑':9500,'mini二節白':9500,
         '固定桌腳':3980,'固定黑腳':3980,'固定白腳':3980,
         'force':23500,'force桌腳':23500,'force四柱桌腳':23500,'force四柱黑腳':23500,'force四柱白腳':23500,}
顏色_list={'':0,'纖維板':0,'菸草橡木':0,'雪白柚木':0,'密西根楓木':0,'北歐白橡木':0,'典雅胡桃木':0,'歐風胡桃木':0,'黑':0,'白':0,'電競':0,'加拿大楓木':0, 
         '蜂巢板':1500,'黑雲岩':1500,'白雲岩':1500,'泥灰岩':1500,'安藤清水模':1500,'台灣柚木':1500,'北海道榆木':1500,'維吉尼亞楓木':1500,'安德森雪松':1500,'哥倫比亞胡桃':1500,}
形狀_list={'':0,'四方前上斜':0,'弧度上斜':0,'四方全平':0,'前凹':800,'後凹':800,'前凹+後凹':800,'小圓弧後凹':800,'四方前上斜+後凹':800,'四方前上斜+小圓弧後凹':800,'四角導圓':800}

while True:
    print('\n')
    try:
       報價 =input('請輸入mini或prime或force或訂製或製材所或下單:')
    except Exception as e:
        print('輸入錯誤，原因:', e)
    else:

        if 報價 == 'mini':
            try:
                桌寬 = float(input('請輸入桌寬:'))
                桌深 = float(input('桌深請輸入60:'))
                桌腳 = input('請輸入桌腳:')
                顏色 = input('請輸入顏色:')
                形狀 = input('請輸入形狀:')

                mini升降桌price = float(mini價目表[桌腳][桌寬])
                顏色price=顏色_list[顏色]
                形狀price=形狀_list[形狀]   
            except Exception as e:
                print('輸入錯誤，原因:', e) 
            else:
                total=mini升降桌price+顏色price+形狀price

                print('和您報價')
                print('%3.0f*%2.0f%s(%s)+%s=%5.0f'%(桌寬,桌深,顏色,形狀,桌腳,total))    
                
        elif 報價 == 'prime':
            try:
                桌寬 = float(input('請輸入桌寬:'))
                桌深 = float(input('請輸入桌深:'))
                桌腳=input('請輸入桌腳:')
                顏色=input('請輸入顏色:')
                形狀=input('請輸入形狀:')
    
                prime升降桌price = float(prime價目表[桌腳][桌寬][桌深])
                顏色price=顏色_list[顏色]
                形狀price=形狀_list[形狀]
            except Exception as e:
                print('輸入錯誤，原因:', e)
            else:
                total=prime升降桌price+顏色price+形狀price
                if 45<桌深<=60:
                    桌腳=桌腳+'(60)'

                print('和您報價')
                print('%3.0f*%2.0f%s(%s)+%s=%5.0f'%(桌寬,桌深,顏色,形狀,桌腳,total))
    
        elif 報價 == 'force':
            try:
                桌寬 = float(input('請輸入桌寬:'))
                桌深 = float(input('請輸入桌深:'))
                桌腳=input('請輸入桌腳:')
                顏色=input('請輸入顏色:')
                形狀=input('請輸入形狀:')
    
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
                顏色=input('請輸入顏色:')
                形狀=input('請輸入形狀:')

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
                if price == 0:
                    print('請修改訂製尺寸，桌深最小50')
                elif 桌寬<110 and (桌腳=='prime三節' or 桌腳=='prime二節' or 桌腳=='prime三節黑' or 桌腳=='prime二節黑'or 桌腳=='prime三節白' or 桌腳=='prime三節灰'or 桌腳=='prime二節灰') and 桌深<57.5:
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
        
                elif 桌寬<110 and (桌腳=='prime三節' or 桌腳=='prime二節' or 桌腳=='prime三節黑' or 桌腳=='prime二節黑'or 桌腳=='prime三節白' or 桌腳=='prime二節白' or 桌腳=='prime三節灰'or 桌腳=='prime二節灰') and 桌深>=57.7:
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
               
                elif 桌深<68 and (桌腳=='prime三節' or 桌腳=='prime二節' or 桌腳=='prime三節黑' or 桌腳=='prime二節黑'or 桌腳=='prime三節白' or 桌腳=='prime二節白' or 桌腳=='prime三節灰'or 桌腳=='prime二節灰'):
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
                                
                elif 桌深<60 and (桌腳=='mini三節' or 桌腳=='mini二節' or 桌腳=='mini三節白' or 桌腳=='mini二節白' or 桌腳=='mini三節黑' or 桌腳=='mini二節黑'):
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

                if  (木種 =='栓木脂接' or 木種 =='栓木直拼' or 木種 =='白橡木脂接' or 木種 =='白橡木直拼'):
                    try:
                        桌寬=float(input('請輸入桌寬:'))
                        桌深=float(input('請輸入桌深:'))
                        厚度=float(input('請輸入厚度(若白橡木直拼3.3請輸入3.5):'))
                        桌腳=input('請輸入桌腳:') 
                        桌腳price=桌腳_list[桌腳] 

                    except Exception as e:
                        print('輸入錯誤，原因:', e)
                    else:
                        if 桌深 >= 81 and 厚度<=2.7:
                            print('無法製作!!桌深厚度81cm以上，厚度需為2.7以上')
                        else:
                            桌板price=round(桌寬*桌深*厚度/2700*木種成本單價list[木種]*木種對客單價乘積list[木種],-2)
                            total=桌板price+桌腳price
                            print('和您報價')
                            print('(製材所)-訂%3.0f*%3.0f*%1.1f%s+%s=%5.0f' % (桌寬,桌深,厚度,木種,桌腳,total))
    
                        print('\n''備註:')
                        print('1. 1.以上金額含運(不含宜花東地區)、不含安裝')
                        print('2. 訂製約45工作天')
                        print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                elif 木種 == '琥珀木':
                    try:
                        桌寬=float(input('請輸入桌寬:'))
                        桌深=float(input('請輸入桌深:'))
                        厚度=float(input('厚度請輸入4.5:'))
                        桌腳=input('請輸入桌腳:') 
             
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
                    except Exception as e:
                        print('輸入錯誤，原因:', e)
                    else:
                        total=桌板price+桌腳price
                        print('和您報價')
                        print('(製材所)-訂%3.0f*%3.0f*%1.1f%s+%s=%5.0f' % (桌寬,桌深,厚度,木種,桌腳,total))
    
                        print('\n''備註:')
                        print('1. 1.以上金額含運(不含宜花東地區)、不含安裝')
                        print('2. 訂製約45工作天')
                        print('3. 製程為工廠製作時間，不包含後續的配送與安裝，實際配送日以通知為主')
                else:
                    print('木種輸入錯誤!!')

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
                custom_weight = int(input('請輸入重量：'))
                custom_url_code = input('請輸入連結編碼: ')
            except Exception as e:
                print('輸入錯誤，原因:', e)
            else:

                payload = {
                    "product": {
                        "title_translations": {
                            "zh-hant": f"【{custom_name}】-{custom_product}"
                        },
                        "category_ids":[
                            "5f3e6edecbccfd003c3fc43f","6620f0d9de0f7400174633cd"
                        ],
                        "price": 99999,  # 原價，可視情況調整
                        "price_sale": custom_price,
                        "unlimited_quantity": True,
                        "sku": "CD01",
                        "weight": custom_weight,
                        "status": "hidden",
                        "retail_status":"draft",
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