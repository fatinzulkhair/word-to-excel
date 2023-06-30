import streamlit as st
from docx import Document
import pandas as pd
from io import BytesIO
import re

st.set_page_config(layout="wide")  # Ukuran layar

st.markdown(
    '<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">',
    unsafe_allow_html=True,
)

st.markdown(
    """
<nav class="navbar navbar-dark bg-dark">
  <a class="navbar-brand" href="#">
    <img src="https://i.pinimg.com/280x280_RS/1f/38/89/1f3889be1f5de92051e8692216d54df5.jpg" width="30" height="30" class="d-inline-block align-top" alt="">
    FatinZulkhair
  </a>
  <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarSupportedContent" aria-controls="navbarSupportedContent" aria-expanded="false" aria-label="Toggle navigation">
    <span class="navbar-toggler-icon"></span>
  </button>

  <div class="collapse navbar-collapse" id="navbarSupportedContent">
    <ul class="navbar-nav mr-auto">
      <li class="nav-item active">
        <a class="nav-link" href="#">Home <span class="sr-only">(current)</span></a>
      </li>
      <li class="nav-item">
        <a class="nav-link" href="#">Link</a>
    </ul>
  </div>
</nav>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """# **Word to Excel App**
A simple app to convert *multiple choice answers* into *table* .
"""
)

st.header("**Select File**")

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


# st.set_page_config(layout="wide")  # Ukuran layar

st.title("Word to Excel")  # Judul

File_word = st.file_uploader("Upload the word file", accept_multiple_files=False)

if File_word is not None:
    table = word_to_excel(File_word).rename_axis("No.").reset_index()
    name_before_extension = re.match("(^.+)(\.\D+)", File_word.name).group(1)
    download_button(table, file_name=f"{name_before_extension}.xlsx")

st.info(
    "Credit: Created by [FatinZulkhair](https://id.linkedin.com/in/m-fatin-zulkhair-yusuf-855427211)"
)

st.markdown(
    """
<script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>
""",
    unsafe_allow_html=True,
)
