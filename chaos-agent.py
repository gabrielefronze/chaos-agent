from docx import Document
from random import shuffle

from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.slider import Slider
from kivy.core.window import Window
from kivy.properties import NumericProperty


class FileSet:
    def __init__(self, filename="test.docx", size=1):
        self.original_file_name = filename
        self.original_file = Document(filename)

        self.shuffled_files_names = []
        for i in range(0, size):
            filename = self.original_file_name.replace(".docx", "_{}.docx".format(i))
            self.shuffled_files_names.append(filename)

        self.shuffled_files = []
        for fn in self.shuffled_files_names:
            self.original_file.save(fn)
            self.shuffled_files.append(Document(fn))

    def shuffle(self):

        for index in range(0, len(self.shuffled_files)):
            texts = {}
            was_list = False

            for i, p in enumerate(self.shuffled_files[index].paragraphs):
                if p.style.name == "List Paragraph":
                    texts[i] = p.text
                    was_list = True
                    # print("{}: That's a list paragraph!".format(i))
                elif was_list or i == len(self.shuffled_files[index].paragraphs)-1:

                    print("mixin'")

                    new_indexes = list(range(list(texts.keys())[0], list(texts.keys())[-1]+1))
                    shuffle(new_indexes)
                    old_indexes = list(texts.keys())
                    shuffle(old_indexes)

                    index_conversion = dict(zip(old_indexes, new_indexes))
                    print(index_conversion)

                    for j, text in texts.items():
                        self.shuffled_files[index].paragraphs[j].text = texts[index_conversion[j]]

                    texts.clear()
                    was_list = False

    def save(self):
        for i, sf in enumerate(self.shuffled_files):
            filename = self.original_file_name.replace(".docx", "_{}.docx".format(i))
            sf.save(filename)


class Container(FloatLayout):
    slider_val = NumericProperty(25)

    # layout
    def __init__(self, *args, **kwargs):
        super(Container, self).__init__(*args, **kwargs)
        Window.clearcolor = (1, 1, 1, 1)

        self.size=(50, 300)
        btn1 = Button(text="Create!", size_hint=(1, .05), pos_hint={'x':.0, 'y':.65})
        btn1.bind(on_press=self.buttonClicked)

        self.txt1 = TextInput(text='',
                              multiline=False,
                              size_hint=(.6, .05),
                              pos_hint={'x': .3, 'y': .85})

        self.txt2 = TextInput(text='',
                              multiline=False,
                              size_hint=(.2, .05),
                              pos_hint={'x': .5, 'y': .75})

        self.lblslider = Label(text="25",
                               size_hint=(1, .05), pos_hint={'x':.0, 'y':.55},
                               color=(0, 0, 0, 1))

        self.slider = Slider(min=1,
                             max=40,
                             value=25,
                             step=1,
                             size_hint=(.6, .05),
                             pos_hint={'x': .3, 'y': .75})
        self.slider.bind(value=self.on_value)

        self.lbl1 = Label(text="Pippo",
                          size_hint=(1, .05),
                          pos_hint={'x': .0, 'y': .55},
                          color=(0, 0, 0, 1))

        self.add_widget(Label(text="Original file name:",
                                size_hint=(.4, .05),
                                pos_hint={'x': .0, 'y': .85},
                                color=(0, 0, 0, 1)))
        self.add_widget(self.txt1)
        self.add_widget(Label(text="Number of output files:",
                                size_hint=(.4, .05),
                                pos_hint={'x': .0, 'y': .75},
                                color=(0, 0, 0, 1)))
        # layout.add_widget(self.txt2)
        self.add_widget(self.slider)
        self.add_widget(self.lbl1)
        self.add_widget(btn1)

    # button click function
    def buttonClicked(self, btn):
        filename = self.txt1.text
        size = int(self.slider.value)

        self.lbl1.text = "Creating {} files from {} original file.".format(size, filename)

        shuffler = FileSet(filename=filename, size=size)
        shuffler.shuffle()
        shuffler.save()

    def on_value(self, instance, value):
        print(value)
        self.lbl1.text = str(value)


class KivyApp(App):
    def build(self):
        return Container()

if __name__ == '__main__':
    KivyApp().run()
