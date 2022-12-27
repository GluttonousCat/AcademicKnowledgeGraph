class WOSdata:
    # 将每一个论文
    def __init__(self):
        self.author = []
        self.nation = []  # 作者所在国籍，C1                              ***
        self.references = []
        self.WC = []  # WOS关键词，WC                                    ***
        self.year = 0
        self.referSOname = []  # txt中的期刊名简写
        self.name = ''  # 写入数据库的TL                                  ***
        self.SO = ''  # 本篇论文所在期刊，已经不用
        self.SOtype = ''  # 本篇论文所在期刊的类型，已经不用
        self.referSO = []  # 写入数据库的引用期刊名，已经不用
        self.referSOtype = []  # 写入数据库的引用期刊类型，与上一个对应        ***
        self.referSOtypeDict = {}  # 写入数据库的引用期刊类型，与上一个对应        ***
        self.keyword = []  # 论文主题词，DE                                ***
        self.org = []  # 作者所在机构，C1                                  ***
        self.tag = []  # 论文标签，又textrank算法得到
        self.abstract = ''  # 论文摘要


class Author:
    def __init__(self):
        self.authonName = '-'
        self.authonOrg = '-'
        self.authorNation = '-'
