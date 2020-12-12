import random
import stylecloud

with open('wordcloud-text.txt', 'r') as reader:
    lines = reader.readlines()
words = []
for line in lines:
    word, weight = line.strip().split(' ')
    words.extend([word] * int(weight))

random.shuffle(words)

stylecloud.gen_stylecloud(
    background_color='#89b6a5',
    colors=['#555555'],
    icon_name='fas fa-square',
    invert_mask=True,
    size=(1200, 627),
    text=' '.join(words),
)
