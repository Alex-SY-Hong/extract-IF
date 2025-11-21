# 用于提取pdf格式的论文对应的影响因子以及期刊名。

## Structure

### 期刊名的获取
提取metadata的subjects，并用正则表达式过滤出期刊名。

### IF的获取
为了更好的稳健性，直接使用Clarivate提供的影响因子文件（xlsx格式）

## 使用

```bash
git clone https://github.com/Alex-SY-Hong/extract-IF.git
cd metadata-and-IF
pip install -r requirements.txt
```
实际上依赖的库只有PyPDF2和pandas，手动下载也可。

另外，我在仓库里附了一个（以CC-BY-4.0协议开放获取）的论文，可供测试使用。

## 问题
不同的文件，metadata格式实际上参差不齐。比如，有文章的subjects字段是摘要，而非
期刊信息。面对参差不齐的metadata，正则表达式这些“传统”工具不很稳健，好在数量
不大的情况下手动处理例外比较容易。

理论上AI可以改善parse的效率。比如说，把首页pdf转成图片，上传给GLM-4.5V之类
的模型提取期刊名（如果有版权等问题，可以考虑本地部署llava等）。有待验证。

## 碎碎念
我要求claude尽量把功能封装为函数，它就大量地使用Maybe（return xx or none)，
属于学到函数式的精髓了。

没有彻底封装，部分设置采取了硬编码。主要是考虑到python程序不太适合分发
给技术小白，把图形化页面和解释器打包到一起会让这个简单程序变得非常臃肿。
