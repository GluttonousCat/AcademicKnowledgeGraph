# -*- coding: utf-8 -*-
import scrapy
import re
from scrapy.http import Request
from scrapy.http import FormRequest
import time
from bs4 import BeautifulSoup
import os
import sys


# wos导出的时候有些批次可能会比实际少一两条，不是本程序的BUG
class WosAdvancedQuerySpiderSpider(scrapy.Spider):
    name = 'wos_advanced_query_spider'
    allowed_domains = ['webofknowledge.com']

    # 提取URL中的QID所需要的正则表达式
    qid_pattern = r'qid=(\d+)&'

    def __init__(self, query=None, output_path='../output', document_type='Article', output_format='fieldtagged', gui=None, sid="", *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.query = query
        self.output_path_prefix = output_path
        self.document_type = document_type
        self.output_format = output_format
        self.gui = gui
        self.sid = sid
        self.downloaded = 0

        if query is None:
            print('请指定检索式')
            sys.exit(-1)
        if output_path is None:
            print('请指定有效的输出路径')
            sys.exit(-1)
        if self.sid != '':
            self.start_urls = [f'http://apps.webofknowledge.com/UA_AdvancedSearch_input.do?SID={self.sid}&product=UA&search_mode=AdvancedSearch']

            print('使用给定的SID：', self.sid)

    def parse(self, response):
        """
        获取SID并提交高级搜索请求，将高级搜索请求返回给parse_result_entry处理
        每次搜索都更换一次SID

        :param response:
        :return:
        """
        sid = self.sid

        adv_search_url = 'http://apps.webofknowledge.com/WOS_AdvancedSearch.do'

        # 检索式，目前设定为期刊，稍作修改可以爬取任意检索式
        print(self.query)
        print(adv_search_url)
        # 将这一个高级搜索请求yield给parse_result_entry，内容为检索历史记录，包含检索结果的入口
        # 同时通过meta参数为下一个处理函数传递sid、journal_name等有用信息
        query_form = {
            "product": "WOS",
            "search_mode": "AdvancedSearch",
            "SID": sid,
            "input_invalid_notice": "Search Error: Please enter a search term.",
            "input_invalid_notice_limits": " <br/>Note: Fields displayed in scrolling boxes must be combined with at least one other search field.",
            "action": "search",
            "replaceSetId": "",
            "goToPageLoc": "SearchHistoryTableBanner",
            "value(input1)": self.query,
            "value(searchOp)": "search",
            "value(select2)": "LA",
            "value(input2)": "",
            "value(select3)": "DT",
            "value(input3)": self.document_type,
            "value(limitCount)": "14",
            "limitStatus": "collapsed",
            "ss_lemmatization": "On",
            "ss_spellchecking": "Suggest",
            "SinceLastVisit_UTC": "",
            "SinceLastVisit_DATE": "",
            "period": "Range Selection",
            "range": "ALL",
            "startYear": "2020",
            "endYear": time.strftime('%Y'),
            # "editions": self.db_list,
            # "editions": ["SCI", "SSCI", "AHCI", "ISTP", "ISSHP", "ESCI", "CCR", "IC"],
            "update_back2search_link_param": "yes",
            "ss_query_language": "",
            "rs_sort_by": "PY.D;LD.D;SO.A;AU.A",
        }
        yield FormRequest(adv_search_url, method='POST', formdata=query_form, dont_filter=True,
                          callback=self.parse_result_entry,
                          meta={'sid': sid, 'query': self.query})


    def parse_result_entry(self, response):
        """
        找到高级检索结果入口链接，交给parse_results处理
        同时还要记录下QID
        :param response:
        :return:
        """
        print(response)
        sid = response.meta['sid']
        query = response.meta['query']

        # 通过bs4解析html找到检索结果的入口
        # BeautifulSoup 支持多种元素定位方式。
        soup = BeautifulSoup(response.text, 'lxml')
        entry_url = soup.find('a', attrs={'title': 'Click to view the results'}).get('href')
        entry_url = 'http://apps.webofknowledge.com' + entry_url
        print(entry_url)

        # 找到入口url中的QID，存放起来以供下一步处理函数使用
        pattern = re.compile(self.qid_pattern)
        result = re.search(pattern, entry_url)
        if result is not None:
            qid = result.group(1)
            print('提取得到qid：', result.group(1))
        else:
            qid = None
            print('qid提取失败')
            exit(-1)

        # yield一个Request给parse_result，让它去处理搜索结果页面，同时用meta传递有用参数
        yield Request(entry_url, callback=self.parse_results,
                      meta={'sid': sid, 'query': query, 'qid': qid})

    def parse_results(self, response):
        print(response)
        sid = response.meta['sid']
        query = response.meta['query']
        qid = response.meta['qid']

        # 通过bs4获取页面结果数字，得到需要分批爬取的批次数
        soup = BeautifulSoup(response.text, 'lxml')
        paper_num = int(soup.find('span', attrs={'id': 'footer_formatted_count'}).get_text().replace(',', ''))
        totalstart = 0

        if paper_num-totalstart > 1000000:
            thisend = 1000000+totalstart
        else:
            thisend = paper_num
        print("大于2.5万数据本次下载取前2.5万")
        span = 500
        iter_num = (thisend-totalstart) // span + 2

        # 对每一批次的结果进行导出500一批
        print('共有{}条文献需要下载'.format(paper_num))
        for i in range(1, iter_num):
            end = i * span+totalstart
            start = (i - 1) * span + 1+totalstart
            if end > paper_num+totalstart:
                end = paper_num+totalstart
            print('正在下载第 {} 到第 {} 条文献'.format(start, end))
            output_form = {
                "selectedIds": "",
                "displayCitedRefs": "true",
                "displayTimesCited": "true",
                "displayUsageInfo": "true",
                "viewType": "summary",
                "product": "WOS",
                "rurl": response.url,
                "mark_id": "WOS",
                "colName": "WOS",
                "search_mode": "AdvancedSearch",
                "locale": "en_US",
                "view_name": "WOS-summary",
                "sortBy": "PY.D;LD.D;SO.A;VL.D;PG.A;AU.A",
                "mode": "OpenOutputService",
                "qid": str(qid),
                "SID": str(sid),
                "format": "saveToFile",
                "filters": "HIGHLY_CITED HOT_PAPER OPEN_ACCESS PMID USAGEIND AUTHORSIDENTIFIERS ACCESSION_NUM FUNDING SUBJECT_CATEGORY JCR_CATEGORY LANG IDS PAGEC SABBR CITREFC ISSN PUBINFO KEYWORDS CITTIMES ADDRS CONFERENCE_SPONSORS DOCTYPE CITREF ABSTRACT CONFERENCE_INFO SOURCE TITLE AUTHORS  ",
                "mark_to": str(end),
                "mark_from": str(start),
                "queryNatural": str(query),
                "count_new_items_marked": "0",
                "use_two_ets": "false",
                "IncitesEntitled": "no",
                "value(record_select_type)": "range",
                "markFrom": str(start),
                "markTo": str(end),
                "fields_selection": "HIGHLY_CITED HOT_PAPER OPEN_ACCESS PMID USAGEIND AUTHORSIDENTIFIERS ACCESSION_NUM FUNDING SUBJECT_CATEGORY JCR_CATEGORY LANG IDS PAGEC SABBR CITREFC ISSN PUBINFO KEYWORDS CITTIMES ADDRS CONFERENCE_SPONSORS DOCTYPE CITREF ABSTRACT CONFERENCE_INFO SOURCE TITLE AUTHORS  ",
                "save_options": "fieldtagged"
            }

            # 将下载地址yield一个FormRequest给download_result函数，传递有用参数
            output_url = 'http://apps.webofknowledge.com/OutboundService.do?action=go&&save_options=othersoftware&'
            yield FormRequest(output_url, method='POST', formdata=output_form, dont_filter=True,
                               callback=self.download_result,
                               meta={'sid': sid, 'query': query, 'qid': qid,
                                     'start': start, 'end': end, 'paper_num': paper_num})

    def download_result(self, response):
        file_postfix = 'txt'
        sid = response.meta['sid']
        query = response.meta['query']
        qid = response.meta['qid']
        start = response.meta['start']
        end = response.meta['end']
        paper_num = response.meta['paper_num']
        print(response)
        # 按日期时间保存文件
        filename = self.output_path_prefix + '/advanced_query/{}/{}.{}'.format(self.query[3:], str(start) + '-' + str(end), file_postfix)
        os.makedirs(os.path.dirname(filename), exist_ok=True)
        text = response.text

        if self.output_format == 'bibtex':
            text = text.replace('Early Access Date', 'Early-Access-Date').replace('Early Access Year', 'Early-Access-Year')
        with open(filename, 'w', encoding='utf-8') as file:
            file.write(text)

        print('--成功下载第 {} 到第 {} 条文献--'.format(start, end))



        self.downloaded += end-start+1
        if self.gui is not None:
            self.gui.ui.progressBarDownload.setValue(self.downloaded/paper_num * 100)

