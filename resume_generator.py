import streamlit as st
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import os
import openai  # Import OpenAI Python client



def generate_resume(name, email, phone, linkedin, summary, programming_languages, libraries, business_intelligence,
                    data_engineering, other_platforms, profile, company_name, jd, degree, university, certifications,
                    additional_skills):
    doc = Document()

    # Set narrow margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Mm(10)
        section.bottom_margin = Mm(10)
        section.left_margin = Mm(10)
        section.right_margin = Mm(10)

    # Add a table for the entire resume
    table = doc.add_table(rows=18, cols=1)  # Increased rows to accommodate additional content
    table.style = 'Table Grid'

    # Name
    cell = table.cell(0, 0)
    cell.text = name
    cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    cell.paragraphs[0].runs[0].font.bold = True

    # Contact information
    cell = table.cell(1, 0)
    cell.text = f" {email}    {phone}   {linkedin}"
    cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Summary
    cell = table.cell(2, 0)
    cell.text = 'Summary'
    cell.paragraphs[0].runs[0].font.bold = True
    cell = table.cell(3, 0)
    cell.text = f"""I am an experienced professional with a background in {profile} at {company_name}.They have expertise in {', '.join(programming_languages)}, {', '.join(business_intelligence)},{', '.join(data_engineering)}, and other platforms. {summary}"""

    # Skills
    cell = table.cell(4, 0)
    cell.text = 'Skills'
    cell.paragraphs[0].runs[0].font.bold = True
    cell = table.cell(5, 0)
    cell.text = f"• Programming Languages:{','.join(programming_languages)}\n• Libraries:{','.join(libraries)}\n• Business Intelligence:{','.join(business_intelligence)}\n• Data Engineering:{','.join(data_engineering)}\n• Other Platforms:{other_platforms}"

    # Company and Profile Info
    cell = table.cell(6, 0)
    cell.text = 'Experience'
    cell.paragraphs[0].runs[0].font.bold = True
    cell = table.cell(7, 0)
    cell.text = f"{profile}\n{company_name}"

    # Experience Description
    cell = table.cell(8, 0)
    # cell.paragraphs[0].runs[0].font.bold = True
    cell.text = '\n'.join(['• ' + j for j in jd.split("\n")])

    # Education
    cell = table.cell(9, 0)
    cell.text = 'Education'
    cell.paragraphs[0].runs[0].font.bold = True
    cell = table.cell(10, 0)
    cell.text = f'{degree}\n{university}'

    # Certifications
    cell = table.cell(11, 0)
    cell.text = 'Certifications'
    cell.paragraphs[0].runs[0].font.bold = True
    cell = table.cell(12, 0)
    cell.text = '\n'.join(['• ' + c for c in certifications.split("\n")])

    # Additional Skills
    cell = table.cell(13, 0)
    cell.text = 'Additional Skills'
    cell.paragraphs[0].runs[0].font.bold = True
    cell = table.cell(14, 0)
    cell.text = '\n'.join(['• ' + ad for ad in additional_skills.split("\n")])

    # Specify the absolute path to save the document
    file_path = os.path.join(os.getcwd(), 'resume.docx')
    doc.save(file_path)
    return file_path


def main():
    st.title('ATS Friendly Resume Generator')

    # Input fields
    name = st.text_input('Name',placeholder='Enter your Full Name')
    email = st.text_input('Email',placeholder='Enter a new email for interview email purposes only')
    phone = st.text_input('Phone',placeholder='Enter your Phone No')
    linkedin = st.text_input('LinkedIn',placeholder='Enter your linkedin url page')
    summary = st.text_area('Summary', placeholder='Write a Brief Summary of your past experience')
    programming_languages = st.multiselect(label='Programming Languages',
                                           options=['Python', 'SQL', 'MS Sql Server', 'MySQL'])
    libraries = st.multiselect(label='Libraries', options=['Numpy', 'Pandas', 'Matplotlib', 'Seaborn', 'Plotly', 'OpenCV'])
    business_intelligence = st.multiselect(label='Business Intelligence',
                                           options=['Power BI', 'Tableau', 'Zoho', 'Quicksight', 'Google Studio',
                                                    'Excel'])
    data_engineering = st.multiselect(label='Data Engineering',
                                      options=['SQL', 'ETL', 'SSIS', 'Data Warehousing', 'Azure Data Factory',
                                               'Data Mining', 'Data Wrangling'])
    other_platforms = st.text_input('Other Platforms',placeholder='Enter other platforms')
    profile = st.text_input('Previous Profile',placeholder='Enter your profile/job role')
    company_name = st.text_input('Company Name',placeholder='Enter your Company Name')
    jd = st.text_area('Explain Job Description',placeholder='Explain what was your job role and what tools you used')
    degree = st.text_input('Education',placeholder='Enter your Highest Qualification')
    university = st.text_input('University',placeholder='Enter University Name')
    certifications = st.text_area('Certifications',placeholder='List all the certifications in each line.')
    additional_skills = st.text_area('Additional Skills',placeholder='List all the additional skills in each line')

    # Generate resume
    if st.button('Generate Resume'):
        if name and email and phone and linkedin and summary and programming_languages and libraries and \
                business_intelligence and data_engineering and other_platforms and company_name and profile and jd and \
                degree and university and certifications and additional_skills:
            file_path = generate_resume(name, email, phone, linkedin, summary, programming_languages, libraries,
                                        business_intelligence, data_engineering, other_platforms, profile,
                                        company_name, jd, degree, university, certifications, additional_skills)
            st.success('Resume generated successfully!')
            st.download_button(
                label="Download your resume",
                data=open(file_path, 'rb').read(),
                file_name='resume.docx',
                mime='application/octet-stream'
            )


if __name__ == '__main__':
    main()
