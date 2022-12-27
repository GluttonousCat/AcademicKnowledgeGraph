class lunwen_node:
    def __init__(self):
        self.label = '论文'
        self.ID = 0
        self.name = ''
        self.publish_year = 0
        self.author = []
        self.keyword = []
        self.WOS_keyword = []
        self.refer_category = []

    def writeIntoTxt(self):
        with open("paper_Node.txt","a+",encoding='utf-8') as fp:
            fp.write('{'+'Label:' + self.label + ', ID:' + str(self.ID) + ', title:' + self.name + ', publish_year:'
                    '' + str(self.publish_year) + ', author:' + str(self.author) + ', keyword:'
                    '' + str(self.keyword) + ', WOS_keyword:' + str(self.WOS_keyword) + ', refer_category:'
                    '' + str(self.refer_category) + '}' + '\n'
                     )


class zuozhe_node:
    def __init__(self):
        self.label = '作者'
        self.ID = 0
        self.name = ''
        self.organization = ''
        self.nation = ''

    def writeIntoTxt(self):
        with open("author_Node.txt", "a+", encoding='utf-8') as fp:
            fp.write('{' + 'Label:' + self.label + ', ID:' + str(self.ID) + ', Name:' + str(self.name) + ', Organization:' + str(self.organization) +', Nation:' + str(self.nation) +'}' + '\n')


class jigou_node:
    def __init__(self):
        self.label = ''
        self.ID = 0
        self.name = ''

    def writeIntoTxt(self):
        with open("organization_Node.txt", "a+", encoding='utf-8') as fp:
            fp.write('{' + 'Label:' + self.label + ', ID:' + str(self.ID) + ', Name:' + str(self.name)  +'}' + '\n')


class guojia_node:
    def __init__(self):
        self.label = ''
        self.ID = 0
        self.name = ''

    def writeIntoTxt(self):
        with open("nation_Node.txt", "a+", encoding='utf-8') as fp:
            fp.write('{' + 'Label:' + self.label + ', ID:' + str(self.ID) + ', Name:' + str(self.name) + '}' + '\n')


class xueke_node:
    def __init__(self):
        self.label = ''
        self.ID = 0
        self.name = ''

    def writeIntoTxt(self):
        with open("category_Node.txt", "a+", encoding='utf-8') as fp:
            fp.write('{' + 'Label:' + self.label + ', ID:' + str(self.ID) + ', Name:' + str(self.name) + '}' + '\n')