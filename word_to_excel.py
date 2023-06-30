import streamlit as st
from docx import Document
import pandas as pd
from io import BytesIO
import re

st.markdown('<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">', unsafe_allow_html=True)

st.markdown("""
<nav class="navbar fixed-top navbar-expand-lg navbar-dark" style="background-color: #3498DB;">
  <a class="navbar-brand" href="https://youtube.com/dataprofessor" target="_blank">Data Professor</a>
  <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarNav" aria-controls="navbarNav" aria-expanded="false" aria-label="Toggle navigation">
    <span class="navbar-toggler-icon"></span>
  </button>
  <div class="collapse navbar-collapse" id="navbarNav">
    <ul class="navbar-nav">
      <li class="nav-item active">
        <a class="nav-link disabled" href="#">Home <span class="sr-only">(current)</span></a>
      </li>
      <li class="nav-item">
        <a class="nav-link" href="https://youtube.com/dataprofessor" target="_blank">YouTube</a>
      </li>
      <li class="nav-item">
        <a class="nav-link" href="https://twitter.com/thedataprof" target="_blank">Twitter</a>
      </li>
    </ul>
  </div>
</nav>
""", unsafe_allow_html=True)

def word_to_excel(nama_file):
    doc = Document(nama_file)
    columns_soal = ["Soal-soal", "A", "B", "C", "D", "E"]
    T_soal_soal = pd.DataFrame(columns=columns_soal)

    count = 0
    count_n = 0

    numbers = []
    for i in doc.paragraphs:
        if re.match(f"^\d+[.)]", i.text):
            star_of_answer = True
            T_soal_soal.loc[count, "Soal-soal"] = i.text[3:]
            count += 1
            count_n += 1
            numbers.append(i.text[3:])
        elif re.match("^\D+[.)]", i.text):
            tar_of_answer = False
            if re.match("^[Aa][.)]", i.text):
                T_soal_soal.loc[(count - 1), "A"] = i.text[3:]
            elif re.match("^[Bb][.)]", i.text):
                T_soal_soal.loc[(count - 1), "B"] = i.text[3:]
            elif re.match("^[Cc][.)]", i.text):
                T_soal_soal.loc[(count - 1), "C"] = i.text[3:]
            elif re.match("^[Dd][.)]", i.text):
                T_soal_soal.loc[(count - 1), "D"] = i.text[3:]
            elif re.match("^[Ee][.)]", i.text):
                T_soal_soal.loc[(count - 1), "E"] = i.text[3:]
        elif re.match("^[Jj][Aa][Ww][Aa][Bb][Aa][Nn]", i.text) or re.match(
            "^[Jj][Aa][Ww][Aa][Bb]", i.text
        ):
            # print(i.text)
            pass
        elif re.match("^\D+", i.text):
            try:
                if star_of_answer == True:
                    T_soal_soal.loc[count - 1, "Soal-soal"] = (
                        str(numbers[(count_n - 1)]) + "\n" + i.text
                    )
                    numbers.append(T_soal_soal.loc[count - 1, "Soal-soal"] + " ")
                    count_n += 1
            except:
                pass

    T_soal_soal.index = T_soal_soal.index + 1
    return T_soal_soal


def download_button(df, file_name):
    # Create an in-memory Excel file
    excel_file = BytesIO()
    with pd.ExcelWriter(excel_file, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")

    # Set the file pointer to the beginning of the file
    excel_file.seek(0)

    # Create a download button
    st.download_button(
        label="Download Excel",
        data=excel_file,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


st.set_page_config(layout="wide")  # Ukuran layar

st.title("Word to Excel")  # Judul

File_word = st.file_uploader("Upload the word file", accept_multiple_files=False)

if File_word is not None:
    table = word_to_excel(File_word).rename_axis("No.").reset_index()
    name_before_extension = re.match("(^.+)(\.\D+)", File_word.name).group(1)
    download_button(table, file_name=f"{name_before_extension}.xlsx")
