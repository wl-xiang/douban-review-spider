import pandas as pd
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import jieba

# 读取Excel文件
df = pd.read_excel('./comments.xlsx', usecols=['Index', 'Review'])

# 对评论文本进行分词和清洗
def process_text(text):
    seg_list = jieba.cut(text, cut_all=False)
    return " ".join(seg_list)

# 加载停用词文件
with open('stopwords.txt', 'r', encoding='utf-8') as file:
    stopwords = set(file.read().split())

# 清洗文本并生成词云
wordcloud_texts = []

for index, row in df.iterrows():
    comment = row['Review']
    processed_comment = process_text(comment)
    words = [word for word in processed_comment.split() if word not in stopwords]
    cleaned_comment = ' '.join(words)
    wordcloud_texts.append(cleaned_comment)

# 合并所有清洗后的评论文本
all_comments = ' '.join(wordcloud_texts)

# 设置词云参数
wordcloud = WordCloud(width=1920, 
                      height=1080, 
                      font_path='simhei.ttf', 
                      background_color='white', 
                      max_words=2000, 
                      stopwords=stopwords,
                      min_font_size=10, 
                      max_font_size=300
            ).generate(all_comments)

# 显示词云
# plt.figure(figsize=(25, 12.5))
plt.figure(figsize=(25, 15))
plt.imshow(wordcloud, interpolation='bilinear')
plt.axis('off')
plt.show()

# 保存词云图像
# plt.savefig('./douban_reviews_wordcloud.png', format='png')
plt.close()