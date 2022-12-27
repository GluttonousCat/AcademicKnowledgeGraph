import os
import time
import py2neo
import EsiData
from NEO4J import NEO4J_DB
from WOSdata import WOSdata, Author
from textrank4zh import TextRank4Keyword, TextRank4Sentence


# 功能：给定路径一级路径，返回二级路径下所有txt的内容
#  exp a--b--1.txt
#     |  |-2.txt
#     |-c--3.txt
# [in]一级路径（exp中的a）
# [out]所有文档内容之和
def getAllTxtData(base_path, graph_entey, stop_words):
# def getAllTxtData(base_path):
    f_woskeyword = open(r'D:\output\WC_data.txt', 'w', encoding='utf-8')
    files = os.listdir(base_path)
    g_WOSdata = WOSdata()
    written = 0
    start = time.time()
    for path in files:
        full_path = os.path.join(base_path, path)
        # print(full_path)
        second_files = os.listdir(full_path)
        for second_path in second_files:
            all_path = os.path.join(full_path, second_path)
            with open(all_path, "r", encoding='utf-8') as fp:
                # 从文本中提取有效信息
                while True:
                    line = fp.readline()
                    # 结束符
                    if line.find('EF') == 0:
                        break

                    # 作者
                    if line.find('AF ') == 0:
                        newAuthor = Author()
                        newAuthor.authonName = line[3:].replace('\n','')
                        g_WOSdata.author.append(newAuthor)
                        line = fp.readline()
                        while line.find('TI ') != 0:
                            newAuthor = Author()
                            newAuthor.authonName = line[3:].replace('\n', '')
                            g_WOSdata.author.append(newAuthor)
                            line = fp.readline()

                    # 论文名
                    if line.find('TI ') == 0:
                        g_WOSdata.name = line[3:].replace('\n',' ').replace('\"','\\"')
                        line = fp.readline()
                        while line.find('SO ') != 0:
                            g_WOSdata.name += line[3:].replace('\n', ' ').replace('\"', '\\"')
                            line = fp.readline()

                    # 主题词
                    if line.find('DE ') == 0:
                        zhutici = line[3:].replace('\n', '').replace('\"', '\\"')
                        line = fp.readline()
                        while line[0] == ' ':
                            zhutici = zhutici + ' ' + line[3:].replace('\n', '').replace('\"', '\\"')
                            line = fp.readline()
                        g_WOSdata.keyword = zhutici.split('; ')
                        index = 0
                        while index < len(g_WOSdata.keyword):
                            g_WOSdata.keyword[index] = g_WOSdata.keyword[index].lower()
                            index += 1

                    # 论文标签，由textrank算法得到
                    if line.find('AB ') == 0:
                        all_abstract = line[3:].replace('\n', '').replace('\"', '\\"')
                        line = fp.readline()
                        # 若第二行仍有文字，继续加入全摘要
                        if line[2] != ' ':
                            all_abstract += line[3:].replace('\n', '').replace('\"', '\\"')
                            line = fp.readline()
                        # 由textrank得到标签tag，由另一个算法得到一句话摘要
                        tr4w = TextRank4Keyword(stop_words_file=stop_words)
                        tr4w.analyze(text=all_abstract.replace('.', '\n'), window=5, lower=True)
                        for item in tr4w.get_keywords():
                            g_WOSdata.tag.append(item['word'])
                        tr4s = TextRank4Sentence()
                        tr4s.analyze(text=all_abstract.replace('.', '\n'), lower=True)
                        for item in tr4s.get_key_sentences(num=1):
                            g_WOSdata.abstract = item['sentence']

                    # 作者所在国家、作者所在机构
                    if line.find('C1 ') == 0:
                        # 国家
                        if checkNation(line) != -1:
                            g_WOSdata.nation.append(checkNation(line))
                        # 所在机构
                        if line.find('] ') != -1:
                            splitedStr =line[(line.find('] ')+1):].replace('\n', '').replace('\"', '\\"').split(', ')
                            g_WOSdata.org.append(splitedStr[0].strip())
                        else:
                            splitedStr = line[3:].replace('\n','').replace('\"','\\"').split(', ')
                            g_WOSdata.org.append(splitedStr[0].strip())
                        # 写入作者信息(国籍, 机构)
                        # 先遍历从AF中获取的作者和作者名字，通过“[]”获取作者名字与同一行的机构和国籍匹配
                        splitedStrList = line[(line.find('[') + 1):(line.find('] '))].split('; ')
                        for authorname in splitedStrList:
                            writtenFlag = False
                            for AF in g_WOSdata.author:
                                # 名字在AF存在
                                if authorname == AF.authonName:
                                    if len(g_WOSdata.nation) > 0 and len(g_WOSdata.org) > 0:
                                        AF.authonOrg = g_WOSdata.org[-1]
                                        AF.authorNation = g_WOSdata.nation[-1]
                                        writtenFlag = True
                            # 名字在AF不存在
                            if writtenFlag == False:
                                newAuthor = Author()
                                newAuthor.authonName = authorname
                                if len(g_WOSdata.nation) > 0 and len(g_WOSdata.org) > 0:
                                    newAuthor.authorNation = g_WOSdata.nation[-1]
                                    newAuthor.authonOrg = g_WOSdata.org[-1]

                        line = fp.readline()
                        while line[0] == ' ':
                            if checkNation(line) != -1:
                                g_WOSdata.nation.append(checkNation(line))
                            if line.find('] ') != -1:
                                splitedStr = line[(line.find('] ') + 1):].replace('\n', '').replace('\"', '\\"').split(', ')
                                g_WOSdata.org.append(splitedStr[0].strip())
                            else:
                                splitedStr = line[3:].replace('\n', '').replace('\"', '\\"').split(',')
                                g_WOSdata.org.append(splitedStr[0].strip())
                            splitedStrList = line[(line.find('[') + 1):(line.find('] '))].split('; ')
                            for authorname in splitedStrList:
                                writtenFlag = False
                                for AF in g_WOSdata.author:
                                    # 名字在AF存在
                                    if authorname == AF.authonName:
                                        if len(g_WOSdata.nation) > 0 and len(g_WOSdata.org) > 0:
                                            AF.authonOrg = g_WOSdata.org[-1]
                                            AF.authorNation = g_WOSdata.nation[-1]
                                            writtenFlag = True
                                # 名字在AF不存在
                                if writtenFlag == False:
                                    newAuthor = Author()
                                    newAuthor.authonName = authorname
                                    if len(g_WOSdata.nation) > 0 and len(g_WOSdata.org) > 0:
                                        newAuthor.authorNation = g_WOSdata.nation[-1]
                                        newAuthor.authonOrg = g_WOSdata.org[-1]
                            line = fp.readline()

                    # 若没有C1字段,则查找RP字段
                    if len(g_WOSdata.nation) == 0 and len(g_WOSdata.org) == 0:
                        if line.find('RP ') == 0:
                            # 国家
                            if checkNation(line) != -1:
                                g_WOSdata.nation.append(checkNation(line))
                            # 所在机构
                            if line.find('] ') != -1:
                                splitedStr = line[(line.find('] ') + 1):].replace('\n', '').replace('\"', '\\"').split(', ')
                                g_WOSdata.org.append(splitedStr[0].strip())
                            else:
                                splitedStr = line[3:].replace('\n', '').replace('\"', '\\"').split(',')
                                g_WOSdata.org.append(splitedStr[0].strip())
                            line = fp.readline()

                    # 引用期刊
                    if line.find('CR ') == 0:
                        refer = line[3:].replace('\n', '').split(', ')
                        if len(refer) >= 3:
                            g_WOSdata.referSOname.append(refer[2])
                        line = fp.readline()
                        while line.find('NR ') != 0:
                            refer = line[3:].replace('\n', '').split(', ')
                            if len(refer) >= 3:
                                g_WOSdata.referSOname.append(refer[2])
                            line = fp.readline()
                        # 将CR值改为full title 并提取类别,同时提取SO的类别
                        ex = EsiData.excelData()
                        if ex.searchByCR(g_WOSdata.SO)[0] != -1:
                            g_WOSdata.SOtype = ex.searchByCR(g_WOSdata.SO)[1]
                        for CR in g_WOSdata.referSOname:
                            ret = ex.searchByCR(CR)
                            if ret[0] != -1 :
                                # g_WOSdata.referSO.append(ret[0])
                                g_WOSdata.referSOtype.append(ret[1])

                    # 引用期刊数量
                    if line.find('NR') == 0:
                        NR = line[3:].replace('\n','')

                    # 出版年份
                    if line.find('PY') == 0:
                        g_WOSdata.year = line[3:].replace('\n', '')

                    # WOS类别
                    if line.find('WC ') == 0:
                        WOStype = line[3:].replace('\n', '')
                        line = fp.readline()
                        while line[0] == ' ':
                            WOStype = WOStype + ' ' + line[3:].replace('\n', '')
                            line = fp.readline()
                        g_WOSdata.WC = WOStype.split('; ')

                    # 单个论文结束
                    if line.find('ER') == 0:
                        # 计数器
                        written += 1
                        if written % 10 == 0:
                            end = time.time()
                            print(str(written)+"已写入,用时", end-start)
                            start = time.time()
                        # 将引文的期刊类型列表改为字典：key:类别 val:数量
                        g_WOSdata.referSOtypeDict = referSOtypeListToDict(g_WOSdata.referSOtype)

                        # 去重
                        g_WOSdata.referSOtype = simplyList(g_WOSdata.referSOtype)
                        g_WOSdata.nation = simplyList(g_WOSdata.nation)
                        g_WOSdata.org = simplyList(g_WOSdata.org)
                        g_WOSdata.keyword = simplyList(g_WOSdata.keyword)

                        # 将值写入数据库
                        db = NEO4J_DB()
                        db.graph = graph_entey
                        # print(g_WOSdata.name)
                        # db.writeToDataBase(g_WOSdata)
                        if True:
                            try:
                                db.writeToDataBase(g_WOSdata)
                            except py2neo.database.DatabaseError:
                                print('error'+g_WOSdata.name)
                            except py2neo.database.ClientError:
                                print('error' + g_WOSdata.name)
                            else:
                                pass

                        '''
                        测试用
                        '''
                        # print(g_WOSdata.tag,g_WOSdata.nation)

                        g_WOSdata.author.clear()
                        g_WOSdata.WC.clear()
                        g_WOSdata.nation.clear()
                        g_WOSdata.org.clear()
                        g_WOSdata.keyword.clear()
                        g_WOSdata.referSOname.clear()
                        g_WOSdata.referSO.clear()
                        g_WOSdata.referSOtype.clear()
                        g_WOSdata.tag.clear()
# end getAllTxtData


def simplyList(List):
    newList = []
    for item in List:
        if item not in newList:
            newList.append(item)
    return newList


def writeToTxt(f2,g_WOSdata : WOSdata):
    if 'China' in g_WOSdata.nation:
        f2.write('China\t')
        for keyword in g_WOSdata.WC:
            f2.write(keyword+'\t')
        f2.write('\n')
    if 'USA' in g_WOSdata.nation:
        f2.write('USA\t')
        for keyword in g_WOSdata.WC:
            f2.write(keyword + '\t')
        f2.write('\n')
    pass


def referSOtypeListToDict(referSOtype):
    referSOtypeDict = {}
    for types in referSOtype:
        if types not in referSOtypeDict.keys():
            referSOtypeDict[types] = 1
        else :
            referSOtypeDict[types] += 1
    return referSOtypeDict


def checkNation(str):
    nationList =[
        'USA',
        'England',
        'Abkhazia',
        'Afghanistan',
        'Albania',
        'Algeria',
        'Andorra',
        'Angola',
        'Antigua and Barbuda',
        'Argentina',
        'Armenia',
        'Australia',
        'Austria',
        'Azerbaijan',
        'Commonwealth oftheBahamas',
        'Bahrain',
        'Bangladesh',
        'Barbados',
        'Belarus',
        'Belgium',
        'Belize',
        'Benin',
        'Bhutan',
        'Bolivia',
        'Bosnia and Herzegovina',
        'Botswana',
        'Brazil',
        'Brunei',
        'Bulgaria',
        'Burkina Faso',
        'BurundiCambodia',
        'Cameroon',
        'Canada',
        'Cape Verde',
        'Catalen',
        'Central African Republic',
        'Chad',
        'Chile',
        'China',
        'Colombia',
        'Comoros',
        'Congo (Brazzaville)',
        'Congo (Kinshasa)',
        'Cook Islands',
        'Costa Rica',
        'Côte d\'Ivoire',
        'Croatia',
        'Cuba',
        'Cyprus',
        'Czech Republic',
        'Denmark',
        'Djibouti',
        'Donetsk People\'s Republic',
        'Dominica',
        'Dominican Republic',
        'Ecuador',
        'Egypt',
        'El Salvador',
        'Equatorial Guinea',
        'Eritrea',
        'Estonia',
        'Ethiopia',
        'Fiji',
        'Finland',
        'France',
        'Gabon',
        'Gambia',
        'Georgia',
        'Germany',
        'Ghana',
        'Greece',
        'Grenada',
        'Guatemala',
        'Guinea',
        'Guinea-Bissau',
        'Guyana',
        'Haiti',
        'Honduras',
        'Hungary',
        'Iceland',
        'India',
        'Indonesia',
        'Iran',
        'Iraq',
        'Ireland',
        'Israel',
        'Italy',
        'Jamaica',
        'Japan',
        'Jordan',
        'Kazakhstan',
        'Kenya',
        'Kiribati',
        'South Korea',
        'Kosovo',
        'Kuwait',
        'Kyrgyzstan',
        'Laos',
        'Latvia',
        'Lebanon',
        'Lesotho',
        'Liberia',
        'Libya',
        'Liechtenstein',
        'Lithuania',
        'Luxembourg',
        'Madagascar',
        'Malawi',
        'Malaysia',
        'Maldives',
        'Maltese Knights',
        'Mali',
        'Malta',
        'Marshall Islands',
        'Mauritania',
        'Mauritius',
        'Mexico',
        'Micronesia',
        'Moldova',
        'Monaco',
        'Mongolia',
        'Montenegro',
        'Morocco',
        'Mozambique',
        'Myanmar',
        'Nagorno-Karabakh',
        'Namibia',
        'Nauru',
        'Nepal',
        'Netherlands',
        'New Zealand',
        'Nicaragua',
        'Niger',
        'Nigeria',
        'Niue',
        'Northern Cyprus',
        'North Macedonia',
        'Norway',
        'Oman',
        'Pakistan',
        'Palau',
        'Palestine',
        'Panama',
        'Papua New Guinea',
        'Paraguay',
        'People\'s Republic of Korea',
        'Peru',
        'Philippines',
        'Poland',
        'Portugal',
        'Pridnestrovie',
        'Puntland',
        'Qatar',
        'Romania',
        'Russia',
        'Rwanda',
        'Saint Christopher and Nevis',
        'Saint Lucia',
        'Saint Vincent and the Grenadines',
        'Samoa',
        'San Marino',
        'São Tomé and Príncipe',
        'Saudi Arabia',
        'Senegal',
        'Serbia',
        'Seychelles',
        'Sierra Leone',
        'Singapore',
        'Slovakia',
        'Slovenia',
        'Solomon Islands',
        'Somali',
        'Somaliland',
        'South Africa',
        'South Ossetia',
        'South Sudan',
        'Spain',
        'Sri Lanka',
        'Sudan',
        'Suriname',
        'Swaziland',
        'Sweden',
        'Switzerland',
        'Syria',
        'Tajikistan',
        'Tanzania',
        'Thailand',
        'Timor-Leste',
        'Togo',
        'Tonga',
        'Trinidad and Tobago',
        'Tunisia',
        'Turkey',
        'Turkmenistan',
        'Tuvalu',
        'Uganda',
        'Ukraine',
        'United Arab Emirates',
        'United Kingdom',
        'United States',
        'Uruguay',
        'Uzbekistan',
        'Vanuatu',
        'Vatican city(the Holy see)',
        'Venezuela',
        'Vietnam',
        'Western Sahara',
        'Yemen',
        'Zambia',
        'Zimbabwe'
    ]
    for nation in nationList:
        if str.find(nation) != -1:
            return nation
    return -1


if False:
    base_path = r'D:\output\written_over'
    getAllTxtData(base_path)


