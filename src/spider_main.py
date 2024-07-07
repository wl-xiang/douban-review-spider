import requests
import time
import pandas as pd
from bs4 import BeautifulSoup

# 豆瓣电影评论主页URL
main_page_url_base = 'https://movie.douban.com/subject/36081094/reviews?start='

# 请求头，模拟浏览器访问
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
}


excel_filename = './douban_comments3.xlsx'
test_filename = './douban_comments_test.xlsx'
idx = 1
max_text_length = 32767  # Excel单元格最大字符数限制
scaned_pages = 0

# 读取已有的Excel文件，如果文件存在
try:
    df_existing = pd.read_excel(excel_filename)
    idx = df_existing['Index'].max() + 1
except FileNotFoundError:
    df_existing = pd.DataFrame(columns=['Index', 'Review'])

review_data = []

# 遍历所有评论页
for i in range(0, 7000, 20):  # 每页20条评论
    url = main_page_url_base + str(i)

    # 发送HTTP GET请求
    response = requests.get(url, headers=headers)

    _page_review_urls = []

    # 确保请求成功
    if response.status_code == 200:
        # 使用BeautifulSoup解析HTML内容
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 查找所有评论链接
        review_links = soup.find_all('a', href=True)
        
        # 过滤出包含review的链接并打印
        for link in review_links:
            href = link['href']
            if 'review/' in href and '#comments' not in href and 'best/' not in href:
                _page_review_urls.append(href)
                # print(len(_page_review_urls))
                # print(href)
        print('--- got {} comments at page {} ---'.format(len(_page_review_urls), int(i / 20 + 1)))
        
    else:
        # 打印状态码
        print('[code {}]Failed to retrieve the webpage.'.format(response.status_code))

    # 获取评论内容
    for review_url in _page_review_urls:
        time.sleep(1)   # 避免被检测为爬虫，设置适当的等待时间

        # 发送HTTP GET请求
        review_response = requests.get(review_url, headers=headers)

        # 检查请求是否成功
        if review_response.status_code == 200:
            # 解析网页内容
            soup = BeautifulSoup(review_response.text, 'html.parser')
            
            # 假设评论内容在某个特定的class或id中，这里需要根据实际网页结构调整
            reviews = soup.find_all('div', class_='review-content')
            
            # 提取评论文本
            comments = [review.get_text() for review in reviews]
            comments_truncated = [comment[:max_text_length - 1] if len(comment) > max_text_length else comment for comment in comments]
            
            # 打印或存储评论内容
            for comment in comments_truncated:
                # # print(comment)
                # if idx % 5 == 0:
                #     print('appending NO.{} comment...'.format(idx))
                # review_data.append((idx, comment))
                # idx += 1
                try:
                    review_data.append((idx, comment))
                    if idx % 5 == 0:
                        print('appending NO.{} comment...'.format(idx))
                    idx += 1
                except Exception as e:
                    print(f"Error appending comment, skipping: {e}")
                    idx += 1
        else:
            # 打印状态码
            print('[code {}]Failed to retrieve the webpage.'.format(review_response.status_code))

    # 先把本页评论数据写入test excel，判断有无写入错误，没有则把本页的评论数据加入到existing_Data
    if review_data:
        df_to_save = pd.DataFrame(review_data, columns=['Index', 'Review'])
        try:
            print(' ===== Testing data writing to excel... ====== ')
            df_to_save.to_excel(test_filename, index=False)
            print(' ===== Testing data written to Excel. ====== ')
            df_existing = pd.concat([df_existing, df_to_save], ignore_index=True)
            print(' ===== Reviews at page {} have been concatted to df_existing. ====== ', int(i / 20 + 1))
            
        except Exception as e:
            print(f"An error occurred while writing to Excel: {e}")
            print('[error] Skip reviews at page {}', (int(i / 20 + 1)))
    
    review_data = []
    print('[info] current len of df_existing:', len(df_existing))

    # 定期落盘，例如每爬取100条评论
    scaned_pages += 1
    if scaned_pages % 5 == 0:
        print('[info] Saving df_existing to excel...')
        try:
            df_existing.to_excel(excel_filename, index=False)
            print('[info] df_existing saved to excel.')
        except Exception as e:
            print(f"An error occurred while writing to Excel: {e}")

    # 等待一段时间，避免被服务器屏蔽IP
    time.sleep(2)

# # 将剩余的评论数据写入Excel
# if review_data:
#     df_to_save = pd.DataFrame(review_data, columns=['Index', 'Review'])
#     df_existing = pd.concat([df_existing, df_to_save], ignore_index=True)
#     df_existing.to_excel(excel_filename, index=False)
#     print('Final data written to Excel.')

print('Done.')