import re

text = "哈哈哈(你是谁)（我是谁）"
pattern = r'[\(\（][^()（）]*[\)\）]'
result = re.sub(pattern, '', text)
print(result)
