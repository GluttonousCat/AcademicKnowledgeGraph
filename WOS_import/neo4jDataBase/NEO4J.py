from py2neo import Graph
from WOSdata import WOSdata


class NEO4J_DB:
    # 已经创建过的期刊类型节点列表
    createdTypeNodeList = []
    # 已经创建过的期刊节点列表
    createdSONodeList = []
    graph = Graph('http://localhost:7474', username='neo4j', password='123456')

    def __init__(self):
        pass

    def writeToDataBase(self, g_WOSdata : WOSdata):
        checkTitle = self.graph.run('match (a:论文) where a.name="'+g_WOSdata.name+'" return a')
        if len(checkTitle.data()) != 0:
            # 确保论文标题唯一:写入前查找数据库是否已经存在这个标题
            return

        self.createLunwenNode(g_WOSdata)
        cypher = 'match (a:论文) where a.name="' + g_WOSdata.name + '"\n'
        if True:
            if len(g_WOSdata.keyword) != 0:
                self.create_contain_relation(g_WOSdata, cypher)
            if len(g_WOSdata.org) != 0:
                self.create_its_org_relation(g_WOSdata, cypher)
            if len(g_WOSdata.referSOtype) != 0:
                self.create_refer_relation(g_WOSdata, cypher)
            if len(g_WOSdata.nation) != 0:
                self.create_its_nation_relation(g_WOSdata, cypher)
            if len(g_WOSdata.WC) != 0:
                self.create_its_WC_relation(g_WOSdata, cypher)
            if len(g_WOSdata.author) != 0:
                self.createAuthorNode(g_WOSdata, cypher)
            if len(g_WOSdata.tag) != 0:
                self.create_its_tag_relation(g_WOSdata, cypher)
                self.create_its_abstract_relation(g_WOSdata, cypher)
            if len(g_WOSdata.referSOtypeDict.keys()) != 0:
                self.create_refers_relation(g_WOSdata, cypher)

    def createLunwenNode(self, g_WOSdata: WOSdata):
        print(g_WOSdata.name, g_WOSdata.year)
        cypher = 'merge (lunwen:论文{name:"' + g_WOSdata.name + '",year:' + g_WOSdata.year + '})'
        self.graph.run(cypher)

    # 添加作者节点
    def createAuthorNode(self, g_WOSdata: WOSdata, cypher):
        nodeNameList = 'g'
        index = 0
        while index < len(g_WOSdata.author):
            nodeNameList += str(index)
            cypher += ('create (' + nodeNameList + ':作者{name:"' + g_WOSdata.author[index].authonName + '",nation:"' +
                       g_WOSdata.author[index].authorNation + '",org:"' + g_WOSdata.author[
                           index].authonOrg + '"})' + '\n')
            cypher += ('create (a)-[:its_author]->(' + nodeNameList + ')\n')
            index += 1
        # print(cypher)
        self.graph.run(cypher)

    def create_contain_relation(self, g_WOSdata: WOSdata, cypher):
        nodeNameList = 'g'
        index = 0
        while index < len(g_WOSdata.keyword):
            nodeNameList += str(index)
            cypher += ('merge ('+nodeNameList+':关键词{name:"'+g_WOSdata.keyword[index]+'"})'+'\n')
            cypher += ('create (a)-[:contain]->(' + nodeNameList + ')\n')
            index += 1
        # print(cypher)
        self.graph.run(cypher)
        return cypher

    def create_its_tag_relation(self, g_WOSdata: WOSdata, cypher):
        nodeNameList = 'g'
        index = 0
        while index < len(g_WOSdata.tag):
            nodeNameList += str(index)
            cypher += ('merge ('+nodeNameList+':标签{name:"'+g_WOSdata.tag[index]+'"})'+'\n')
            cypher += ('create (a)-[:its_tag]->(' + nodeNameList + ')\n')
            index +=1
        # print(cypher)
        self.graph.run(cypher)
        return cypher

    def create_its_abstract_relation(self, g_WOSdata: WOSdata, cypher):
        nodeNameList = 'g'
        index = 0

        nodeNameList += str(index)
        cypher += ('merge ('+nodeNameList+':摘要{name:"'+g_WOSdata.abstract+'"})'+'\n')
        cypher += ('create (a)-[:its_abstract]->(' + nodeNameList + ')\n')

        # print(cypher)
        self.graph.run(cypher)
        return cypher

    def create_its_org_relation(self,g_WOSdata: WOSdata, cypher):
        nodeNameList = 'g'
        index = 0
        while index < len(g_WOSdata.org):
            nodeNameList += str(index)
            cypher += ('merge ('+nodeNameList+':机构{name:"'+g_WOSdata.org[index]+'"})'+'\n')
            cypher += ('create (a)-[:its_org]->(' + nodeNameList + ')\n')
            index += 1
        # print(cypher)
        self.graph.run(cypher)
        return cypher

    def create_refer_relation(self, g_WOSdata: WOSdata, cypher):
        nodeNameList = ['lei1', 'lei2', 'lei3', 'lei4', 'lei5', 'lei6', 'lei7', 'lei8', 'lei9', 'lei10', 'lei11',
                        'lei12', 'lei13', 'lei14', 'lei15','lei21', 'lei22', 'lei23', 'lei24', 'lei25', 'lei26',
                        'lei27', 'lei28', 'lei29', 'lei20', 'lei31']
        index = 0
        while index < len(g_WOSdata.referSOtype):
            cypher += ('merge ('+nodeNameList[index]+':论文类别{name:"'+g_WOSdata.referSOtype[index]+'"})'+'\n')
            cypher += ('create (a)-[:refer]->(' + nodeNameList[index] + ')\n')
            index += 1
        # print(cypher)
        self.graph.run(cypher)
        return cypher

    def create_its_WC_relation(self, g_WOSdata: WOSdata, cypher):
        nodeNameList = ['wc1', 'wc2', 'wc3', 'wc4', 'wc5', 'wc6', 'wc7', 'wc8', 'wc9', 'wc10', 'wc11', 'wc12', 'wc13',
                        'wc14', 'wc15']
        index = 0
        while index < len(g_WOSdata.WC):
            cypher += ('merge ('+nodeNameList[index]+':WC{name:"'+g_WOSdata.WC[index]+'"})'+'\n')
            cypher += ('create (a)-[:its_WC]->(' + nodeNameList[index] + ')\n')
            index += 1
        # print(cypher)
        self.graph.run(cypher)
        return cypher

    def create_its_nation_relation(self, g_WOSdata: WOSdata, cypher):
        nodeNameList = 'g'
        index = 0
        while index < len(g_WOSdata.nation):
            nodeNameList += str(index)
            cypher += ('merge ('+nodeNameList+':国籍{name:"'+g_WOSdata.nation[index]+'"})'+'\n')
            cypher += ('create (a)-[:its_nation]->(' + nodeNameList + ')\n')
            index += 1
        # print(cypher)
        self.graph.run(cypher)
        return cypher

    # 加引用期刊类别
    def create_refers_relation(self, g_WOSdata: WOSdata, cypher):
        nodeNameList = 'g'
        keyList = list(g_WOSdata.referSOtypeDict.keys())
        for key in keyList:
            cypher += ('create (' + nodeNameList + ':引用期刊类别{name:"' + key + '", num:'+str(g_WOSdata.referSOtypeDict[key])+'})' + '\n')
            cypher += ('create (a)-[:refers]->(' + nodeNameList + ')\n')
            nodeNameList += '1'
        # print(cypher)
        self.graph.run(cypher)
        return cypher