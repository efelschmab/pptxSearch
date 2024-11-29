import customtkinter as ctk


class GUI():

    def __init__(self) -> None:
        pass

    def button(self, master_frame, button_text, button_command, column, row) -> ctk.CTkButton:

        button = ctk.CTkButton(master=master_frame,
                               corner_radius=10,
                               border_width=0,
                               font=("Roboto", 12),
                               text=button_text,
                               command=button_command,
                               height=40)
        button.grid(row=row, column=column, padx=30, sticky="e")
        return button

    def db_label(self, master_frame, text, column) -> ctk.CTkLabel:

        self.db_text_label = ctk.CTkLabel(master=master_frame,
                                          corner_radius=10,
                                          text=text,
                                          anchor="w",
                                          wraplength=300,
                                          justify="left")
        self.db_text_label.grid(row=0, column=column,
                                sticky="ew", padx=(0, 30))
        return self.db_text_label

    def loglabel(self, master_frame, text, column) -> ctk.CTkLabel:

        self.log_label = ctk.CTkLabel(master=master_frame,
                                      corner_radius=10,
                                      text=text,
                                      anchor="w",
                                      wraplength=270,
                                      justify="left")
        self.log_label.grid(row=2, column=column,
                            sticky="ew", padx=30, columnspan=2)
        return self.log_label

    def container(self, master_frame, row):

        container = ctk.CTkFrame(master=master_frame,
                                 height=70,
                                 fg_color="#242424")
        container.grid(sticky="ew", row=row, column=0,
                       padx=10, pady=10, columnspan=2)

    def entry(self, master_frame, column, row) -> ctk.CTkEntry:

        entry_var = ctk.StringVar(value="enter search query")
        self.entry_field = ctk.CTkEntry(master=master_frame,
                                        textvariable=entry_var,
                                        corner_radius=10,
                                        font=("Roboto", 12),
                                        border_width=0)
        # Textvariable -> no placeholder text
        self.entry_field.bind(
            "<FocusIn>", lambda event: self.entry_field.delete(0, "end"))
        self.entry_field.grid(row=row, column=column,
                              sticky="ew", padx=(0, 40))
        return self.entry_field

    def change_label_text(self, text: str) -> None:
        self.db_text_label.configure(text=text)

    def change_log_label_text(self, text: str) -> None:
        self.log_label.configure(text=text)

    def fetch_query_field(self) -> str:

        query: str = self.entry_field.get()
        if query != "" and len(query) > 2:
            return self.entry_field.get()
        else:
            self.entry_field.delete(0, "end")
            return ""
