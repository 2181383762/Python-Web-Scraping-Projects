项目名称：Book Scraper

项目简介：
本项目使用 Python 自动采集 Books to Scrape 测试网站上的图书信息，并导出为 Excel 文件。

采集字段：

- Title：图书名称
- Price：图书价格
- Rating：图书评分
- Availability：库存状态
- Product URL：商品详情页链接

运行环境：

- Python 3.11 或兼容版本
- requests
- beautifulsoup4
- openpyxl

运行方法：

1. 打开项目文件夹。
2. 在地址栏输入 cmd 并按回车。
3. 输入 python first_scraper.py 并按回车。
4. 等待程序运行结束。
5. 在项目文件夹中查看 books.xlsx。

输出文件：
books.xlsx

注意事项：

- 本项目用于学习和演示网页数据采集。
- 运行前请确保网络连接正常。
- 如果 Excel 文件正在打开，请先关闭它，再运行程序。