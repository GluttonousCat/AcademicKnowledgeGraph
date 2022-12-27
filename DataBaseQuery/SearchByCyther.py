from py2neo import Graph, Node, Relationship

from Category_Knowledge_Group.knowledgeGroup import lunwen_node, zuozhe_node, jigou_node, guojia_node, xueke_node


class NEO4J_DB:
    graph = Graph('http://localhost:7474', username='neo4j', password='neo4j')

    def getNumberByKeywordEveryYear(self, leftY, rightY, keyword):
        cypher = 'match (a:论文)-->(b:关键词) where b.name="' + keyword + '" and a.year>='+str(leftY)+'' \
                    'and a.year<='+str(rightY)+'return a.year as 年份,count(*) as 数量 order by a.year'
        # print(cypher)
        ret = self.graph.run(cypher)
        data = ret.data()
        numByYear = []
        year = int(leftY)
        while year <= int(rightY):
            numByYear.append(0)
            year += 1
        for idx in data:
            numByYear[int(idx['年份'])-int(leftY)] = int(idx['数量'])
        ret = {keyword: numByYear}
        # print(ret)
        return ret

    def getKeywodsName(self):
        cypher = 'match (a:`关键词`) return a.name as keyword,count(*) order by a.name limit 100'
        # print(cypher)
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def getKeywodsNameAndCount(self, limit_num=0):
        """
        得到关键词名称与数量
        :param limit_num:
        :return:
        """
        #
        cypher = 'match (n:`论文`)-->(a:`关键词`) return a.name as 关键词,count(*) as 数量 order by count(*) desc limit ' + str(limit_num)
        # print(cypher)
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def changeKeyword(self, before='', after=''):
        cypher = 'match (a:`关键词`) where a.name="'+before+'" set a.name="'+after+'"'
        self.graph.run(cypher)

    def countPaperByYear(self, leftY, rightY):
        cypher = 'match (a:`论文`) where a.year>='+str(leftY)+' and a.year<='+str(rightY)+' return count(*) as 数量'
        # print(cypher)
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def getTagByAuthor(self, author):
        cypher = 'match (a:`论文`)-->(b:`作者`),(a:`论文`)-->(c:`标签`) where b.name = "' + author \
                 + '" return c.name as 标签,count(*) as 数量 order by count(*) desc limit 10'
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def getAbstractByAuthor(self, author):
        cypher = 'match (a:`论文`)-->(b:`作者`),(a:`论文`)-->(c:`摘要`) where b.name = "' + author \
                 + '" return c.name as 摘要 limit 10'
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def getAllTag(self, num):
        cypher = 'match (a:`论文`)-->(b:`标签`) return b.name as tag ,count(*) as 数量 order by count(*) desc limit ' + str(num)
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def getWCDataByNation(self, nation):
        cypher = 'match (a:`论文`)-->(b:WC),(a:`论文`)-->(c:`国籍`) where c.name="'+nation+'" return b.name as WOS,count(*) as number order by count(*) desc'
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def getWCDataByNationByYear(self, nation, leftY, rightY):
        cypher = 'match (a:`论文`)-->(b:WC),(a:`论文`)-->(c:`国籍`) where c.name="'+nation+'" and a.year >= '+str(leftY)+'and a.year <=' +str(rightY)+ ' return b.name as WOS,count(*) as number order by count(*) desc'
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def getReferSOType(self):
        cypher = 'match (a:`引用期刊类别`) return a.name as category,count(*) order by count(*) desc'
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    # 某区间全部引文数, Wa
    def getLeibieByYear(self, leftyear, rightyear):
        cypher = 'match (a:论文)-->(b:引用期刊类别) ' +'\n'\
                 'where a.year>= '+str(leftyear)+'and a.year<='+str(rightyear) + '\n'\
                 'return b.name as category,sum(b.num) as number'+'\n' \
                 'order by sum(b.num) desc'
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    # 某国某区间全部引文数Wb+Wi
    def getLeibieByNationByYearOrderByName(self, nation, leftyear, rightyear):
        cypher = 'match (a:论文)-->(b:引用期刊类别),(a:论文)-->(c:国籍)' +'\n'\
                 'where c.name= "'+nation+'" and a.year>='+str(leftyear)+' and a.year<='+str(rightyear)+'\n'+ \
                 'return b.name as category, sum(b.num) as number' +'\n'\
                 'order by b.name'
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def getLeibieByYearOrderByName(self, leftyear, rightyear):
        cypher = 'match (a:论文)-->(b:引用期刊类别) ' +'\n'\
                 'where a.year>= '+str(leftyear)+'and a.year<='+str(rightyear) + '\n'\
                 'return b.name as category,sum(b.num) as number'+'\n' \
                 'order by b.name'
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def getWCDataByYear(self, year):
        cypher = 'match (a:论文)-->(b:WC) ' \
                 'where a.year = ' + str(year) + ' return b.name as subject, count(*) as number ' \
                 'order by count(*)'
        #print(cypher)
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def getYearList(self):
        cypher = 'match (a:论文) return a.year as year,count(a.year) as number order by a.year'
        #print(cypher)
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def searchAllWC(self):
        cypher =    'match (b:WC)' +'\n'\
                    'return b.name, count(*)'+'\n'\
                    'order by count(*)'
        #print(cypher)
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def searchAllKeyword(self, WC, num):
        cypher =    'match (a:论文)-->(b:国籍),(a)-->(c:关键词),(a:论文)-->(d:WC)' +'\n'\
                    'where b.name = "China" and d.name = "'+WC+'"' +'\n'\
                    'return c.name as 关键词,count(*) as 数量' +'\n'\
                    'order by count(*) DESC' +'\n'\
                    'limit '+num
        #print(cypher)
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def searchAllKeywordByYear(self, WC, num, leftY, rightY):
        cypher =    'match (a:论文)-->(b:国籍),(a)-->(c:关键词),(a:论文)-->(d:WC)' +'\n'\
                    'where b.name = "China" and d.name = "'+WC+'"' +'\n'\
                    'and a.year>= '+str(leftY)+'and a.year<='+str(rightY) + '\n'\
                    'return c.name as 关键词,count(*) as 数量' +'\n'\
                    'order by count(*) DESC' +'\n'\
                    'limit '+num
        #print(cypher)
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def searchAllKeywordOnlyByYear(self, num, leftY, rightY):
        cypher =    'match (a:论文)-->(c:关键词)' +'\n'\
                    'where a.year>= '+str(leftY)+'and a.year<='+str(rightY) + '\n'\
                    'return c.name as 关键词,count(*) as 数量' +'\n'\
                    'order by count(*) DESC' +'\n'\
                    'limit '+str(num)
        # print(cypher)
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def searchCoorperatWithChina(self, WC):
        cypher = 'match (a:论文)-->(b:国籍),(a:论文)-->(c:国籍),(a:论文)-->(d:WC)' +'\n'\
                 'where b.name = "China" and d.name = "'+ WC +'"' +'\n'\
                 'return c.name,count(*)' +'\n'\
                 'order by count(*) desc' +'\n'\
                 'limit 5'
        #print(cypher)
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def searchAllOrgan(self, WC, num):
        cypher = 'match (a:论文)-->(b:机构),(a:论文)-->(d:WC)' +'\n'\
                 'where d.name = "'+WC+'"'+'\n'\
                 'return b.name as 机构,count(*) as 数量'+'\n'\
                 'order by count(*) DESC' +'\n'\
                 'limit ' + num
        #print(cypher)
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def searchAllOrganByYear(self, WC, num, leftY, rightY):
        cypher = 'match (a:论文)-->(b:机构),(a:论文)-->(d:WC)' +'\n'\
                 'where d.name = "'+WC+'"'+'\n' \
                 'and a.year>= ' + str(leftY) + 'and a.year<=' + str(rightY) + '\n' \
                 'return b.name as 机构,count(*) as 数量'+'\n'\
                 'order by count(*) DESC' +'\n'\
                 'limit ' + num
        #print(cypher)
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def searchAllCoorperateNation(self, WC, num):
        cypher = 'match (a:论文)-[r]->(b:国籍),(a:论文)-->(d:WC)' +'\n'\
                 'where d.name = "'+WC+'"' +'\n'\
                 'with a,count(r) as cotr' +'\n'\
                 'where cotr>1' +'\n'\
                 'with a' +'\n'\
                 'match (a:论文)-[r]->(b:国籍)' +'\n'\
                 'return b.name as 国家,count(*) as 数量'+'\n'\
                 'order by count(*) DESC' +'\n'\
                 'limit ' + num
        #print(cypher)
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def searchAllCoorperateNationByYear(self, WC, num, leftY, rightY):
        cypher = 'match (a:论文)-[r]->(b:国籍),(a:论文)-->(d:WC)' +'\n'\
                 'where d.name = "'+WC+'"' +'\n' \
                 'and a.year>= ' + str(leftY) + 'and a.year<=' + str(rightY) + '\n' \
                 'with a,count(r) as cotr' +'\n'\
                 'where cotr>1' +'\n'\
                 'with a' +'\n'\
                 'match (a:论文)-[r]->(b:国籍)' +'\n'\
                 'return b.name as 国家,count(*) as 数量'+'\n'\
                 'order by count(*) DESC' +'\n'\
                 'limit ' + num
        #print(cypher)
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def searchCoorperateNationKeyword(self, nation1, nation2, WC, num):
        cypher = 'match (a:论文)-[r]->(b:国籍)' +'\n'\
                 'with a,count(r) as cotr' +'\n'\
                 'where cotr>1' +'\n'\
                 'with a' +'\n'\
                 'match (a:论文)-->(b:国籍),(a:论文)-->(c:国籍),(a:论文)-->(d:关键词),(a:论文)-->(e:WC)' +'\n'\
                 'where b.name = "'+nation1+'" and c.name = "'+nation2+'"and e.name = "'+WC+'"' +'\n'\
                 'return b.name as 国家1,c.name as 国家2,d.name as 关键词,count(*) as 数量' +'\n'\
                 'order by count(*) DESC' +'\n'\
                 'limit ' + num
        #print(cypher)
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def searchCoorperateNationKeywordByYear(self, nation1, nation2, WC, num, leftY, rightY):
        cypher = 'match (a:论文)-[r]->(b:国籍)' +'\n'\
                 'with a,count(r) as cotr' +'\n'\
                 'where cotr>1' +'\n'\
                 'with a' +'\n'\
                 'match (a:论文)-->(b:国籍),(a:论文)-->(c:国籍),(a:论文)-->(d:关键词),(a:论文)-->(e:WC)' +'\n'\
                 'where b.name = "'+nation1+'" and c.name = "'+nation2+'"and e.name = "'+WC+'"' +'\n' \
                 'and a.year>= ' + str(leftY) + 'and a.year<=' + str(rightY) + '\n' \
                 'return b.name as 国家1,c.name as 国家2,d.name as 关键词,count(*) as 数量' +'\n'\
                 'order by count(*) DESC' +'\n'\
                 'limit ' + num
        #print(cypher)
        ret = self.graph.run(cypher)
        data = ret.data()
        return data

    def searchCoorperateNationOrgan(self, nation1, nation2, WC, num):
        cypher = 'match (a:论文)-[r]->(b:国籍)' +'\n'\
                 'with a,count(r) as cotr' +'\n'\
                 'where cotr>1' +'\n'\
                 'with a' +'\n'\
                 'match (a:论文)-->(b:国籍),(a:论文)-->(c:国籍),(a:论文)-->(d:机构),(a:论文)-->(e:WC)' +'\n'\
                 'where b.name = "'+nation1+'" and c.name = "'+nation2+'"and e.name = "'+WC+'"' +'\n'\
                 'return b.name as 国家1,c.name as 国家2,d.name as 机构,count(*) as 数量' +'\n'\
                 'order by count(*) DESC' +'\n'\
                 'limit ' + num
        #print(cypher)
        ret = self.graph.run(cypher)
        data = ret.data()
        return data
        pass

    def searchCoorperateNationOrganByYear(self, nation1, nation2, WC, num, leftY, rightY):
        cypher = 'match (a:论文)-[r]->(b:国籍)' +'\n'\
                 'with a,count(r) as cotr' +'\n'\
                 'where cotr>1' +'\n'\
                 'with a' +'\n'\
                 'match (a:论文)-->(b:国籍),(a:论文)-->(c:国籍),(a:论文)-->(d:机构),(a:论文)-->(e:WC)' +'\n'\
                 'where b.name = "'+nation1+'" and c.name = "'+nation2+'"and e.name = "'+WC+'"' +'\n'\
                 'and a.year>= ' + str(leftY) + 'and a.year<=' + str(rightY) + '\n'\
                 'return b.name as 国家1,c.name as 国家2,d.name as 机构,count(*) as 数量' +'\n'\
                 'order by count(*) DESC' +'\n'\
                 'limit ' + num
        # print(cypher)
        ret = self.graph.run(cypher)
        data = ret.data()
        return data
        pass

    '''
    导出知识图谱相关
    '''
    def export_lunwen_node(self):
        cypher = 'match (a:论文) return a.name limit 1000'
        ret = self.graph.run(cypher)
        allLunwenTitleList = ret.data()
        #print(allLunwenTitleList)
        for index in allLunwenTitleList:
            # label
            Node = lunwen_node()
            Node.label = '论文'
            # title
            title = index['a.name']
            Node.name = title
            # id, pyblishyear
            cypher = 'match (a:论文) where a.name ="'+ title.replace('\"','\\"') +'" return id(a),a.year'
            ret = self.graph.run(cypher)
            data = ret.data()
            Node.ID = int(data[0]['id(a)'])
            Node.publish_year = int(data[0]['a.year'])
            # author
            cypher = 'match (a:论文)-->(b:作者) where a.name ="'+ title.replace('\"','\\"') +'"return b.name'
            ret = self.graph.run(cypher)
            data = ret.data()
            for i in data:
                Node.author.append(i['b.name'])
            # keyword
            cypher = 'match (a:论文)-->(b:关键词) where a.name ="'+ title.replace('\"','\\"') +'"return b.name'
            ret = self.graph.run(cypher)
            data = ret.data()
            for i in data:
                Node.keyword.append(i['b.name'])
            # WOS_keyword
            cypher = 'match (a:论文)-->(b:WC) where a.name ="'+ title.replace('\"','\\"') +'"return b.name'
            ret = self.graph.run(cypher)
            data = ret.data()
            for i in data:
                Node.WOS_keyword.append(i['b.name'])
            # refer_category
            cypher = 'match (a:论文)-->(b:引用期刊类别) where a.name ="'+ title.replace('\"','\\"') +'"return b.name'
            ret = self.graph.run(cypher)
            data = ret.data()
            for i in data:
                Node.refer_category.append(i['b.name'])
            # 写入txt
            Node.writeIntoTxt()

    def export_zuozhe_node(self):
        cypher = 'match (a:作者) return id(a),a.name,a.nation,a.org'
        ret = self.graph.run(cypher)
        data = ret.data()
        for index in data:
            Node = zuozhe_node()
            Node.ID = int(index['id(a)'])
            Node.name = index['a.name']
            Node.nation = index['a.nation']
            Node.organization = index['a.org']
            Node.writeIntoTxt()

    def export_jigou_node(self):
        cypher = 'match (a:机构) return id(a),a.name'
        ret = self.graph.run(cypher)
        data = ret.data()
        for index in data:
            Node = jigou_node()
            Node.ID = int(index['id(a)'])
            Node.name = index['a.name']
            Node.writeIntoTxt()

    def export_guojia_node(self):
        cypher = 'match (a:国籍) return id(a),a.name'
        ret = self.graph.run(cypher)
        data = ret.data()
        for index in data:
            Node = guojia_node()
            Node.ID = int(index['id(a)'])
            Node.name = index['a.name']
            Node.writeIntoTxt()

    def export_xueke_node(self):
        cypher = 'match (a:论文类别) return id(a),a.name'
        ret = self.graph.run(cypher)
        data = ret.data()
        for index in data:
            Node = xueke_node()
            Node.ID = int(index['id(a)'])
            Node.name = index['a.name']
            Node.writeIntoTxt()

    def export_lunwen_zuozhe_rela(self):
        cypher = 'match (a:论文)-->(b:作者) return id(a),id(b) limit 1000'
        ret = self.graph.run(cypher)
        data = ret.data()
        with open("paper_author_link.txt", "a+", encoding='utf-8') as fp:
            for i in data:
                fp.write('{{Label:论文, ID:' + str(i['id(a)']) + '}-->{Label:作者, ID:' + str(i['id(b)']) +  '}}'+ '\n')

    def export_zuozhe_guojia_rela(self):
        cypher = 'match (a:`论文`)-->(b:`作者`),(a:`论文`)-->(c:`国籍`) where c.name=b.nation return id(b),id(c)'
        ret = self.graph.run(cypher)
        data = ret.data()
        with open("author_nation_link.txt", "a+", encoding='utf-8') as fp:
            for i in data:
                fp.write('{{Label:作者, ID:' + str(i['id(a)']) + '}-->{Label:国家, ID:' + str(i['id(b)']) + '}}' + '\n')

    def export_zuozhe_jigou_rela(self):
        cypher = 'match (a:`论文`)-->(b:`作者`),(a:`论文`)-->(c:`机构`) where c.name=b.nation return id(b),id(c)'
        ret = self.graph.run(cypher)
        data = ret.data()
        with open("author_nation_link.txt", "a+", encoding='utf-8') as fp:
            for i in data:
                fp.write('{{Label:作者, ID:' + str(i['id(a)']) + '}-->{Label:机构, ID:' + str(i['id(b)']) + '}}' + '\n')