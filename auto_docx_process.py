import os
import sys
from collections import Counter
from finereport_docx_process import doc_process


def change_to_docx(file_path, filetype):
    doc = word.Documents.Open(file_path + '.' + filetype)
    doc.SaveAs("{}.docx".format(file_path), 12)
    doc.Close()


def main(doc_file, target_file):
    """
    遍历目录，得到文件夹位置
    :return:
    """
    for industry in os.listdir(doc_file):
        I = os.path.join(doc_file, industry)
        if not os.path.isfile(I):

            for year in os.listdir(I):
                Y = os.path.join(I, year)
                if not os.path.isfile(Y):

                    for month in os.listdir(Y):
                        M = os.path.join(Y, month)
                        if not os.path.isfile(M):

                            for province in os.listdir(M):
                                P = os.path.join(M, province)
                                if not os.path.isfile(P):

                                    for doc_type in os.listdir(P):
                                        if doc_type == 'word':
                                            D = os.path.join(P, doc_type)

                                            docs = os.listdir(D)
                                            doc_names = []
                                            for file in docs:
                                                filename, filetype = os.path.splitext(file)
                                                if filetype == '.doc' or filetype == '.rtf':
                                                    change_to_docx(os.path.join(D, filename), filetype)
                                                    doc_names.append(filename[:len(filename) - 2])
                                                elif filetype == '.docx':
                                                    doc_names.append(filename[:len(filename) - 2])
                                                else:
                                                    continue
                                            x = Counter(doc_names).most_common()
                                            for each in x:
                                                if each[1] == 2:
                                                    text_path = os.path.join(D, each[0] + '图文.docx')
                                                    table_path = os.path.join(D, each[0] + '表格.docx')
                                                    target_path = target_file + '/%s/%s/%s/%s/%s' % (
                                                    industry, year, month, province, doc_type)
                                                    if not os.path.exists(target_path):
                                                        os.makedirs(target_path)
                                                    doc_process(text_path, table_path, target_path+'/%s.docx'% each[0]).run()

if __name__ == '__main__':
    if len(sys.argv) == 1:
        doc_file = r'D:/fineReports/software'
        target_file = r'D:/fineReports/processed'
        main(doc_file, target_file)
    elif len(sys.argv) == 4:
        text_path = sys.argv[1]
        table_path = sys.argv[2]
        target_path = sys.argv[3]
        doc_process(text_path, table_path, target_path).run()
    else:
        print("命令行参数有误，请输入为：图文文件路径，表格文件路径，输出文件路径")
