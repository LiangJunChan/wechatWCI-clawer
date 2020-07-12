from baidu_index.utils import test_cookies
from baidu_index import config
from baidu_index import BaiduIndex, ExtendedBaiduIndex
import openpyxl

cookies = """这里填充 cookies 注意引号数量"""

if __name__ == "__main__":
    # 测试cookies是否配置正确
    # True为配置成功，False为配置不成功
    print(test_cookies(cookies))

    keywords = ['张艺凡', "陈卓璇", '希林娜依·高', '赵粤', '王艺瑾', '郑乃馨', '刘些宁','创造营2020']
    # keywords = ['杨超越','孟美岐','吴宣仪','李紫婷','徐梦洁','傅菁','Yamy','赖美云','段奥娟','杨芸晴','张紫宁','创造101']

    # 获取城市代码, 将代码传入area可以获取不同城市的指数, 不传则为全国
    # 媒体指数不能分地区获取
    print(config.PROVINCE_CODE)
    print(config.CITY_CODE)

    # 获取百度搜索指数(地区为山东)
    baidu_index = BaiduIndex(
        keywords=keywords,
        start_date='2020-05-01',
        end_date='2020-07-05',
        # start_date='2018-04-20',
        # end_date='2018-06-24',
        cookies=cookies,
        # area=901
    )
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['name', 'type', 'value', 'date'])
    for index in baidu_index.get_index():
        # ws.cell("index").value = index
        keyword = index["keyword"]
        data = index["date"]
        type = index["type"]
        index2 = index["index"]
        if type == "all":
            ws.append([keyword, "czy2020", index2, data])
            print(index)

    wb.save("test.xlsx")

    # 获取百度媒体指数
    # news_index = ExtendedBaiduIndex(
    #     keywords=keywords,
    #     start_date='2020-05-01',
    #     end_date='2020-07-07',
    #     cookies=cookies,
    #     kind='news'
    # )
    # for index in news_index.get_index():
    #     print(index)

    # # 获取百度咨询指数
    # feed_index = ExtendedBaiduIndex(
    #     keywords=keywords,
    #     start_date='2020-05-01',
    #     end_date='2020-07-07',
    #     cookies=cookies,
    #     kind='feed'
    # )
    # for index in feed_index.get_index():
    #     print(index)
