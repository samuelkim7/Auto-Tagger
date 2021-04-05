from tkinter import *
import pandas as pd
import re


def excel_to_dic(excel_file):
    word_list = pd.read_excel(excel_file)
    word_to_tag = {}

    i = 0
    for col in word_list.columns:
        # skip the columns in the even position
        if i % 2 == 1:
            i += 1
            continue

        # map word to tag
        for j in range(len(word_list[col])):
            try:
                word = word_list.iloc[j, i]
                tag = word_list.iloc[j, i+1]
                # stop if the column ended (word == NaN)
                if type(word) != str:
                    continue

                if word != tag:
                    word_to_tag[word] = '{{Tag' + col + '|[[' + tag + \
                                        '|' + word + ']]}}'
                else:
                    word_to_tag[word] = '{{Tag' + col + '|[[' + word + ']]}}'
            except:
                raise ValueError('excel file style error')
        i += 1
    print(word_to_tag)
    return word_to_tag


def tag(word_to_tag, input_file_name):
    with open(input_file_name, 'r', encoding='UTF8') as input_file, \
            open(input_file_name.split('.')[0] + '_tagged.txt', 'w', encoding='UTF8') as output:
        content = input_file.read()
        for word, tag in word_to_tag.items():
            # find tags that were assigned before
            assigned_tags = re.findall("{{[^}}]*}}", content)

            # replace the whole content if word is not in tags
            word_in_tags = False
            for assigned_tag in assigned_tags:
                if word in assigned_tag:
                    print(":: Word found in assigned tags ::")
                    word_in_tags = True
                    break

            if not word_in_tags:
                print(":: whole content replacement ::")
                content = content.replace(word, tag, 1)
            else:
                print(":: split and concat ::")
                # split the content by tags
                split_by_tags = re.split("{{[^}}]*}}", content)

                # replace word to tag in split_by_tags (only the first occurrence)
                for i in range(len(split_by_tags)):
                    if split_by_tags[i].find(word) != -1:
                        split_by_tags[i] = split_by_tags[i].replace(word, tag, 1)
                        break

                # concatenate tags and split_by_tags
                # split_by_tags has one more element
                print(f'tags: {assigned_tags}, split_by_tags: {split_by_tags}')
                print(f'length of tags: {len(assigned_tags)}, length of split_by_tags: {len(split_by_tags)}')

                # here

        output.write(content)


def open_window():
    def tagger():
        try:
            word_to_tag = excel_to_dic(entry_excel.get())
            tag(word_to_tag, entry_text.get())

            window2 = Tk()
            window2.title('Success')
            label_2 = Label(window2, text='태깅이 완료되었습니다.')
            label_2.grid(row=1, column=0, padx=10, pady=10)
            button_2 = Button(window2, text='닫기', command=window2.destroy)
            button_2.grid(row=1, column=1, padx=10, pady=10)

            window2.mainloop()
        except FileNotFoundError:
            window2 = Tk()
            window2.title('Error')
            label_2 = Label(window2, text='파일명 및 확장자를 확인하세요.')
            label_2.grid(row=1, column=0, padx=10, pady=10)
            button_2 = Button(window2, text='닫기', command=window2.destroy)
            button_2.grid(row=1, column=1, padx=10, pady=10)

            window2.mainloop()
        except Exception as e:
            print(e)

    window = Tk()
    window.title('Auto Tagger v1.0')

    label1 = Label(window, text='텍스트 파일명 (확장자 포함) 입력:')
    label1.grid(row=1, column=0, padx=10, pady=10)
    entry_text = Entry(window)
    entry_text.grid(row=1, column=1, padx=10, pady=10)

    label2 = Label(window, text='리스트 엑셀 파일명 (확장자 포함) 입력:')
    label2.grid(row=2, column=0, padx=10, pady=10)
    entry_excel = Entry(window)
    entry_excel.grid(row=2, column=1, padx=10, pady=10)

    button = Button(window, text='Tagging', command=tagger)
    button.grid(row=3, padx=10, pady=10)

    window.mainloop()


if __name__ == '__main__':
    word_to_tag = excel_to_dic('Tag Excel List v.1.0.xlsx')
    tag(word_to_tag, 'sample_input.txt')
    # open_window()
