import tkinter as tk
from tkinter import filedialog

from datetime import datetime

import requests
import docx
import os
import urllib.parse
import plotly.express as px
import pandas as pd
import plotly.io as pio
pio.renderers.default = "browser"



class GUI:

    def __init__(self):
        self.root = tk.Tk()

        self.root.geometry("500x600")
        self.root.configure(bg = "#e7e7e7")
        self.root.title("OpenAI Essay Detector")

        self.select_folder_button = tk.Button(self.root, text="Select Folder", font=("RobotoRomanLight", 14), command=self.open_folder_dialog)
        self.select_folder_button.place(x=330.0,y=30.0,width=140.0,height=40.0)

        self.select_folder_entry = tk.Entry(self.root)
        self.select_folder_entry.insert(0,"No Folder Selected")
        self.select_folder_entry.place(x=30.0,y=37.5,width=275.0,height=25.0)

        self.run_analysis_button = tk.Button(self.root, text="Run Analysis",font=("RobotoRomanLight", 24), command=self.run_analysis)
        self.run_analysis_button.place(x=30.0,y=100.0,width=440.0,height=70.0)

        self.show_graph_button = tk.Button(self.root, text="Show Graph", font=("RobotoRomanLight", 24), command=self.graph)
        self.show_graph_button.place(x=30.0,y=200.0,width=205.0,height=70.0)

        self.save_results_button = tk.Button(self.root, text="Save Results", font=("RobotoRomanLight", 24), command=self.save_results)
        self.save_results_button.place(x=265.0,y=200.0,width=205.0,height=70.0)

        self.status_label = tk.Label(self.root, text="Program options:", anchor="w")
        self.status_label.place(x=30.0,y=300.0,width=440.0,height=25.0)

        self.gpt2_option_checkbox_var = tk.IntVar()
        self.gpt2_option_checkbox = tk.Checkbutton(self.root, text="Use OpenAI GPT-2 Output Detetector", font=("RobotoRomanLight", 14), variable=self.gpt2_option_checkbox_var)
        self.gpt2_option_checkbox.place(x=30.0,y=325.0)
        self.gpt2_option_checkbox.select()

        self.ai_cheat_check_option_checkbox_var = tk.IntVar()
        self.ai_cheat_check_option_checkbox = tk.Checkbutton(self.root, text="Use AICheatCheck", font=("RobotoRomanLight", 14), variable=self.ai_cheat_check_option_checkbox_var)
        self.ai_cheat_check_option_checkbox.place(x=30.0,y=355.0)
        self.ai_cheat_check_option_checkbox.select()


        self.axis_option_checkbox_var = tk.IntVar()
        self.axis_option_checkbox = tk.Checkbutton(self.root, text="Use logarithmic x axis for graph", font=("RobotoRomanLight", 14), variable=self.axis_option_checkbox_var)
        self.axis_option_checkbox.place(x=30.0,y=385.0)
    
        self.status_label = tk.Label(self.root, text="Program status:", anchor="w")
        self.status_label.place(x=30.0,y=420.0,width=440.0,height=25.0)

        self.status_window = tk.Text(self.root)
        self.status_window.place(x=30.0,y=445.0,width=440.0,height=120.0)

        self.root.mainloop()

    
    def open_folder_dialog(self):
        self.root.withdraw()
        self.select_folder_entry.delete(0,"end")
        try:
            self.select_folder_entry.insert(0,(filedialog.askdirectory()+"/"))
        except TypeError:
            #user hit cancel
            self.select_folder_entry.delete(0,"end")
            self.select_folder_entry.insert(0,"No Folder Selected")
        self.root.deiconify()

    def run_analysis(self):
        self.analysis_results = []
        self.source_file_path = self.select_folder_entry.get()
        try:
            files_for_analysis = [filename for filename in os.listdir(self.source_file_path) if filename.endswith(".docx")]
        except FileNotFoundError:
            self.status_window.delete("0.0","end")
            self.status_window.insert("0.0",("ERROR: Selected directory does not exist."))
            return None
        if len(files_for_analysis)==0:
            #no files for analysis!
            self.status_window.delete("0.0","end")
            self.status_window.insert("0.0",("ERROR: Selected directory does not contain any files with .docx extension."))
            return None
        if not self.ai_cheat_check_option_checkbox_var.get() and not self.gpt2_option_checkbox_var.get():
            self.status_window.delete("0.0","end")
            self.status_window.insert("0.0",("ERROR: You have not selected a model to use. Please select a model."))
            return None

        def get_plaintext(document: docx.Document) -> str: 
            #get plaintext body of document:
            document_plaintext = [paragraph.text for paragraph in document.paragraphs]
            document_plaintext = [paragraph for paragraph in document_plaintext if len(paragraph)>2]
            return document_plaintext

        def get_score(document: list, filename: str) -> dict:
            gpt2_result=None
            aicheatcheck_result=None

            try:
                if self.gpt2_option_checkbox_var.get():
                    #this works weirdly, we basically have to append paragraphs to the url
                    gpt2_url = "https://openai-openai-detector.hf.space/?"
                    for paragraph in document:
                        gpt2_url += urllib.parse.quote(paragraph, safe="")+"=&"
                    gpt2_url = gpt2_url[:-1]    
                    r = requests.get(url=gpt2_url)
                    if r.status_code != 200:
                        #There is an error:
                        self.status_window.delete("0.0","end")0
                        self.status_window.insert("0.0","ERROR: GPT-2 model returned the following error on file"+filename+":\n"+str(r.status_code)+": "+str(r.reason)+".\nPlease remove file from folder and try again. Feel free to report this error to Tom if you think it is a bug.")
                        raise Exception
                    else:
                        gpt2_result = r.json()
                        gpt2_result = gpt2_result['fake_probability']

                if self.ai_cheat_check_option_checkbox_var.get():
                    #normal post request
                    aicheatcheck_url = "https://demo.aicheatcheck.com/api/detect"
                    data = {'text':("\n".join(document))}
                    r = requests.post(url=aicheatcheck_url,json=data)
                    if r.status_code != 200:
                        #There is an error:
                        self.status_window.delete("0.0","end")0
                        self.status_window.insert("0.0","ERROR: AICheatCheck model returned the following error on file"+filename+":\n"+str(r.status_code)+": "+str(r.reason)+".\nPlease remove file from folder and try again. Feel free to report this error to Tom if you think it is a bug.")
                        raise Exception
                    else:
                        aicheatcheck_result = r.json()
                        aicheatcheck_result = aicheatcheck_result['probability_fake']
            except:
                self.status_window.delete("0.0","end")
                self.status_window.insert("0.0","ERROR: There is an error in the returned value from the models. Check that you are connected to the internet, and that the selected documents contain valid text content.")
                raise Exception
            
            if gpt2_result is None and aicheatcheck_result is None:
                self.status_window.delete("0.0","end")
                self.status_window.insert("0.0","ERROR: The selected models have returned no value. Check that you are connected to the internet, and that the selected documents contain valid text content.")
                raise Exception
            else:
                result_dict = {'GPT-2':gpt2_result,'AICheatCheck':aicheatcheck_result}
                return result_dict
        
        def update_window():
            self.status_window.delete("0.0","end")
            self.status_window.insert("0.0",("Processing file " + str(len(self.analysis_results)+1) + "/" + str(len(files_for_analysis))))
            self.root.update()
        

        #loop for analysis
        for filename in files_for_analysis:
            update_window()
            current_essay = docx.Document(self.source_file_path+filename)
            current_essay_plaintext = get_plaintext(current_essay)
            essay_score = get_score(current_essay_plaintext, filename)
            self.analysis_results.append({"Name":filename, "OpenAI-Generated Probability":essay_score})

        self.result_df = pd.json_normalize(self.analysis_results)

        try:
            self.result_df["OpenAI-Generated Probability.GPT-2"] = self.result_df['OpenAI-Generated Probability.GPT-2'].apply(lambda x: round(x*100,3))
        except:
            pass
        try:
            self.result_df["OpenAI-Generated Probability.AICheatCheck"] = self.result_df['OpenAI-Generated Probability.AICheatCheck'].apply(lambda x: round(x*100,3))
        except:
            pass
        self.status_window.insert("end","\n" + str(len(self.analysis_results)) + " files processed.\n")
        self.root.update()

    def graph(self):
        self.graphing_df = self.result_df.sort_values(by=self.result_df.columns[1])
        self.graphing_df = self.graphing_df.melt(id_vars='Name', var_name="Model", value_name="Probability").replace(to_replace="OpenAI-Generated Probability.", value="",regex=True)
        self.results_fig = px.scatter(self.graphing_df, title="Probability of OpenAI Generation in Document", y="Name", x='Probability', color="Model", log_x=self.axis_option_checkbox_var.get())
        self.results_fig.show()

    def save_results(self):
        dt_now = datetime.now().strftime("%Y%m%d%H%M%S")
        self.result_df.to_csv(self.source_file_path + "OpenAI Essay Check Results " + dt_now + ".csv", index=False)
        self.status_window.insert("end","\n" + "File saved at " + self.source_file_path+"OpenAI Essay Check Results " + dt_now + ".csv")




gui = GUI()