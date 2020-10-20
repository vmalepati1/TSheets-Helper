import tkinter as tk

from google_sheets import GoogleSheets
from list_formatting import Runner


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()

        self.formatter = Runner("S.1__288d9b69996781ce0b70679c80ef1cb1de87ff12", "info.json")
        self.formatter.run()

        self.do_notifications = None
        self.spreadsheet_link_entry = None
        self.create_widgets()

    def create_widgets(self):
        canvas = tk.Canvas(self)

        self.do_notifications = tk.BooleanVar()
        do_notifications_button = tk.Checkbutton(canvas, text="Send Notifications", variable=self.do_notifications,
                                                 onvalue=True, offvalue=False)

        spreadsheet_link_label = tk.LabelFrame(canvas, text="Google Spreadsheet Link (Will use default from info.json "
                                                            "if unprovided)")
        self.spreadsheet_link_entry = tk.Entry(spreadsheet_link_label)

        go_button = tk.Button(canvas, text="Run TSheets Helper", command=self.run)

        canvas.pack()
        tk.Label(canvas, text="TSheets Helper").pack(pady=10)
        do_notifications_button.pack(pady=15)
        spreadsheet_link_label.pack(padx=100)
        self.spreadsheet_link_entry.pack(expand=tk.YES, fill=tk.X, padx=10, pady=10)
        go_button.pack(side=tk.BOTTOM, pady=10, padx=10)

    def get_data(self):
        spreadsheet_id = None
        spreadsheet_link = self.spreadsheet_link_entry.get()

        if spreadsheet_link is not "":
            try:
                spreadsheet_id = spreadsheet_link.split("/d/")[1].split("/")[0]
            except IndexError:
                print("Invalid Google Sheets link. Trying as Spreadsheet ID.")
                spreadsheet_id = spreadsheet_link.split("/")[0]

        return self.do_notifications.get(), spreadsheet_id

    def run(self):
        do_notification, spreadsheet_id_arg = self.get_data()
        print("{}, {}".format(do_notification, spreadsheet_id_arg))

        spreadsheet_id = self.formatter.info['spreadsheet_id']

        if spreadsheet_id_arg is not None:
            spreadsheet_id = spreadsheet_id_arg

        sheets = GoogleSheets(spreadsheet_id)
        sheets.send_to_google_sheets(self.formatter)
        sheets.open_spreadsheet()

        if do_notification:
            print("Sent notifications")
            self.formatter.compose_and_send_notifications(sheets)

        exit()


if __name__ == '__main__':
    root = tk.Tk()
    app = Application(master=root)
    app.mainloop()
