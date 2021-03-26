from tkinter import *
import pandas as pd


def excel_to_dic(excel_file):
    word_list = pd.read_excel(excel_file)
    word_to_tag = {}
    for column in word_list:
        for word in word_list[column]:
            word_to_tag[word] = '{{Tag' + column + '|[[' + word + ']]}}'
    return word_to_tag


def tag(word_to_tag, input_file_name):
    with open(input_file_name, 'r', encoding='UTF8') as input_file, \
            open(input_file_name.split('.')[0] + '_tagged.txt', 'w', encoding='UTF8') as output:
        content = input_file.read()
        for word, tag in word_to_tag.items():
            # replace only the first occurrence
            content = content.replace(word, tag, 1)
        output.write(content)


def open_window():
    def main():
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

    button = Button(window, text='Tagging', command=main)
    button.grid(row=3, padx=10, pady=10)

    window.mainloop()


if __name__ == '__main__':
    open_window()
