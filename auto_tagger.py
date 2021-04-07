from tkinter import *
import pandas as pd
import re


def excel_to_dic(excel_file_name):
    word_list = pd.DataFrame()
    try:
        word_list = pd.read_excel(excel_file_name + '.xlsx')
    except FileNotFoundError as e:
        word_list = pd.read_excel(excel_file_name + '.xls')
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
    return word_to_tag


def tag(word_to_tag, input_file_name):
    with open(input_file_name + '.txt', 'r', encoding='UTF8') as input_file, \
            open(input_file_name.split('.')[0] + '_tagged.txt', 'w', encoding='UTF8') as output:
        content = input_file.read()
        for word, tag in word_to_tag.items():
            # find tags that were assigned before
            assigned_tags = re.findall("{{[^}}]*}}", content)

            # check if word is in assigned_tags
            word_in_tags = False
            for assigned_tag in assigned_tags:
                if word in assigned_tag:
                    print(":: Word found in assigned tags ::")
                    word_in_tags = True
                    break

            # replace the whole content if word is not in assigned_tags
            if not word_in_tags:
                print(":: whole content replacement ::")
                content = content.replace(word, tag, 1)
            # split, replace, and concatenate if word is in assigned_tags
            else:
                print(":: split, replace, and concatenate ::")
                # split the content by assigned_tags
                split_by_tags = re.split("{{[^}}]*}}", content)

                # replace word to tag in split_by_tags (only the first occurrence)
                for i in range(len(split_by_tags)):
                    if split_by_tags[i].find(word) != -1:
                        split_by_tags[i] = split_by_tags[i].replace(word, tag, 1)
                        break

                # concatenate assigned_tags and split_by_tags
                # split_by_tags has one more element than assigned_tags
                print(f'tags: {assigned_tags}, split_by_tags: {split_by_tags}')
                print(f'length of tags: {len(assigned_tags)}, length of split_by_tags: {len(split_by_tags)}')
                concat = ''
                for split, assigned_tag in zip(split_by_tags, assigned_tags):
                    print('concat')
                    concat += split + assigned_tag
                concat += split_by_tags[-1]
                content = concat

        output.write(content)


def open_window():
    def main():
        try:
            word_to_tag = excel_to_dic(entry_excel.get())
            tag(word_to_tag, entry_text.get())

            window2 = Tk()
            window2.title('Success')
            label_2 = Label(window2, text='태깅이 완료되었습니다.', **style)
            label_2.grid(row=1, column=0, **grid_style)
            button_2 = Button(window2, text='닫기', command=window2.destroy)
            button_2.grid(row=1, column=1, **grid_style)

            window2.mainloop()
        except FileNotFoundError:
            window2 = Tk()
            window2.title('Error')
            label_2 = Label(window2, text='파일위치 및 파일명을 확인하세요.', **style)
            label_2.grid(row=1, column=0, **grid_style)
            button_2 = Button(window2, text='닫기', command=window2.destroy)
            button_2.grid(row=1, column=1, **grid_style)

            window2.mainloop()
        except ValueError('excel file style error'):
            window2 = Tk()
            window2.title('Error')
            label_2 = Label(window2, text='엑셀 파일의 형식을 확인하세요.', **style)
            label_2.grid(row=1, column=0, **grid_style)
            button_2 = Button(window2, text='닫기', command=window2.destroy)
            button_2.grid(row=1, column=1, **grid_style)

            window2.mainloop()
        except Exception as e:
            print(e)

    window = Tk()
    window.title('Auto Tagger v1.0')
    style = {'font': ('Malgun Gothic', 10)}
    grid_style = {'padx':13, 'pady':13}

    label2 = Label(window, text='텍스트 파일명 입력:', **style)
    label2.grid(row=1, column=0, **grid_style)
    entry_text = Entry(window, width=30)
    entry_text.grid(row=1, column=1, **grid_style)

    label3 = Label(window, text='리스트 엑셀 파일명 입력:', **style)
    label3.grid(row=2, column=0, **grid_style)
    entry_excel = Entry(window, width=30)
    entry_excel.grid(row=2, column=1, **grid_style)

    button = Button(window, text='Tagging', command=main, **style)
    button.grid(row=3, column=1, **grid_style)

    window.mainloop()


if __name__ == '__main__':
    open_window()
