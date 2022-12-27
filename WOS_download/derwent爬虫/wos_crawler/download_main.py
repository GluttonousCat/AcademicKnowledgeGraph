# from gui.main_gui import *
import sys
import time
import scrapy
from scrapy.spiders import crawl
from twisted.internet import reactor, defer
from scrapy.crawler import CrawlerProcess, CrawlerRunner
from scrapy.utils.project import get_project_settings
import settings
from scrapy import cmdline
sys.path.append(r"D:\python\derwent爬虫\wos_crawler\spiders")
from wos_advanced_query_spider import WosAdvancedQuerySpiderSpider


def crawl_by_query(query, output_path='../output', document_type='Article', output_format='bibtex', sid=''):
    cmdline.execute(
        r'scrapy crawl wos_advanced_query_spider -a output_path={} -a output_format={}'.format(output_path, output_format).split()
        + ['-a', 'query={}'.format(query), '-a', 'document_type={}'.format(document_type), '-a', 'sid={}'.format(sid)])


def run_crawl_process(journalListPath, outputPath, SID):
    name = []
    # 将所有下载期刊存入列表
    with open(journalListPath, 'r+', encoding="UTF-8") as filename:
        for line in filename:
            name.append(line.splitlines())
    # 依次下载列表中的期刊
    runner = CrawlerRunner(get_project_settings())

    @defer.inlineCallbacks
    def a():
        i = 0
        while True:
            i += 1
            yield runner.crawl(WosAdvancedQuerySpiderSpider, query='so=({})'.format(name[i - 1][0]),
                               output_path=outputPath, document_type='',
                               output_format='fieldtagged', sid=SID)
            # time.sleep(60)
        reactor.stop()

    a()
    reactor.run()


if False:
    # 按期刊下载
    # crawl_by_journal(journal_list_path=r'C:\Users\Tom\PycharmProjects\wos_crawler\input\journal_list_test.txt',
    #                  output_path=r'E:\wos爬取结果', output_format='fieldtagged', document_type='')

    name = []
    # 将所有下载期刊存入列表
    with open(r'D:\毕业设计\学习内容\1数据收集\未下载期刊.txt', 'r+', encoding="UTF-8") as filename:
        for line in filename:
            name.append(line.splitlines())
    # 依次下载列表中的期刊
    runner = CrawlerRunner(get_project_settings())
    @defer.inlineCallbacks
    def a():
        i = 0
        while True:
            i += 1
            yield runner.crawl(WosAdvancedQuerySpiderSpider, query='so=({})'.format(name[i-1][0]), output_path=r'D:\output', document_type='',
                          output_format='fieldtagged', sid="8BEMLnAR93Rhkg3VhF9")
            # time.sleep(60)
        reactor.stop()
    a()
    reactor.run()


    # 按检索式下载

    # crawl_by_query(query='',
            # output_path='D://output', sid='8CzcJ9Ntth1JInf2Bhy', output_format='fieldtagged', document_type='')
    # 网页如果重置了 ，那么sid也还是会重置的
    # 每次下载 ，现在网页上检索。然后在点击运行。
    # 使用GUI下载
    # crawl_by_gui()
    pass
'''
1 要打开英文的网页
2 不要超过10000条，已修复
3 要人工查询过后 加入sidsci
'''
'''class Request(object_ref):

    def __init__(self, url, callback=None, method='GET', headers=None, body=None,
                 cookies=None, meta=None, encoding='utf-8', priority=0,
                 dont_filter=False, errback=None, flags=None, cb_kwargs=None):
                 
                 # URL爬取网址
                 # callback 回调函数
                 # method 请求方式
                 # headers 请求头
                 # body 网页代码
                 # cookies 浏览痕迹
                 # meta 传递的信息
                 # encoding 编码方式
                 # priority 优先级
                 # dont_filter 去重
                    
                 '''