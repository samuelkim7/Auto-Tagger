from tkinter import *
import pandas as pd
import re


def excel_to_dic(excel_file_name):
    word_list = pd.DataFrame()
    try:
        word_list = pd.read_excel(excel_file_name + '.xlsx')
    except FileNotFoundError:
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
                    word_in_tags = True
                    break

            # replace the whole content if word is not in assigned_tags
            if not word_in_tags:
                content = content.replace(word, tag, 1)
            # split, replace, and concatenate if word is in assigned_tags
            else:
                # split the content by assigned_tags
                split_by_tags = re.split("{{[^}}]*}}", content)

                # replace word to tag in split_by_tags (only the first occurrence)
                for i in range(len(split_by_tags)):
                    if split_by_tags[i].find(word) != -1:
                        split_by_tags[i] = split_by_tags[i].replace(word, tag, 1)
                        break

                # concatenate assigned_tags and split_by_tags
                # split_by_tags has one more element than assigned_tags
                concat = ''
                for split, assigned_tag in zip(split_by_tags, assigned_tags):
                    concat += split + assigned_tag
                concat += split_by_tags[-1]
                content = concat

        output.write(content)


def detag(input_file_name):
    with open(input_file_name + '.txt', 'r', encoding='UTF8') as input_file, \
            open(input_file_name.split('.')[0] + '_detagged.txt', 'w', encoding='UTF8') as output:
        content = input_file.read()
        assigned_tags = re.findall("{{[^}}]*}}", content)
        split_by_tags = re.split("{{[^}}]*}}", content)

        # remove tags and get the words only
        words = []
        for assigned_tag in assigned_tags:
            # word != tag
            if assigned_tag.count('|') == 2:
                word = assigned_tag.split('|')[-1].split(']')[0]
                words.append(word)
                print(assigned_tag, word)
            # word == tag
            elif assigned_tag.count('|') == 1:
                word = assigned_tag.split('[')[-1].split(']')[0]
                words.append(word)
                print(assigned_tag, word)

        # concatenate words and split_by_tags
        concat = ''
        for split, word in zip(split_by_tags, words):
            concat += split + word
        concat += split_by_tags[-1]
        content = concat

        output.write(content)


def open_window():
    def tagger():
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
        except ValueError:
            window2 = Tk()
            window2.title('Error')
            label_2 = Label(window2, text='엑셀 파일의 형식을 확인하세요.', **style)
            label_2.grid(row=1, column=0, **grid_style)
            button_2 = Button(window2, text='닫기', command=window2.destroy)
            button_2.grid(row=1, column=1, **grid_style)

            window2.mainloop()
        except Exception:
            pass

    def detagger():
        try:
            detag(entry_text.get())

            window2 = Tk()
            window2.title('Success')
            label_2 = Label(window2, text='태그 제거가 완료되었습니다.', **style)
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
        except Exception:
            pass

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

    button1 = Button(window, text='Tagging', command=tagger, **style)
    button1.grid(row=3, column=1, padx=13, pady=5)

    button2 = Button(window, text='Detagging', command=detagger, **style)
    button2.grid(row=4, column=1, **grid_style)

    window.mainloop()


if __name__ == '__main__':
    open_window()
