import streamlit as st
from docx import Document
from docx.shared import Mm, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import os


def generate_resume(name, city, area_name, zipcode, email, phone, linkedin, summary,
                    programming_languages, libraries, business_intelligence, data_engineering,
                    big_data, profile, company_name, start_date, end_date, is_current_job, jd,
                    degree, university, certifications, additional_skills, statistical_methods,
                    data_collection, database_management, cloud_platforms, machine_learning):
    """Generates an ATS-friendly resume with a visually appealing layout."""
    doc = Document()

    # Set narrow margins for a sleek design
    for section in doc.sections:
        section.top_margin = Mm(10)
        section.bottom_margin = Mm(10)
        section.left_margin = Mm(10)
        section.right_margin = Mm(10)

    # Helper function to add and style cells
    def add_and_style_cell(text, bold=False, centered=False, font_size=12):
        cell = table.add_row().cells[0]  # Add a new row and get the first cell
        paragraph = cell.paragraphs[0]
        run = paragraph.add_run(text)
        run.bold = bold
        run.font.size = Pt(font_size)
        if centered:
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Initially, create a table with a single row for the header
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'

    # Add resume content
    add_and_style_cell(name, bold=True, centered=True, font_size=14)
    add_and_style_cell('Data Analyst', centered=True)
    contact_info = f"{city}, {area_name}, {zipcode} | {email} | {phone} | {linkedin}"
    add_and_style_cell(contact_info, centered=True)
    add_and_style_cell('Summary', bold=True)
    add_and_style_cell(summary)

    # Dynamic content from user input (skills, experience, etc.)
    add_and_style_cell('Technical Skills', bold=True)
    skills = [programming_languages, libraries, business_intelligence, data_engineering,
              big_data, statistical_methods, data_collection, database_management,
              cloud_platforms, machine_learning]
    for skill_set in skills:
        if skill_set:  # Check if the skill set is not empty
            skill_text = ', '.join(skill_set)
            add_and_style_cell(skill_text)

    add_and_style_cell('Professional Experience', bold=True)
    experience_text = f"{profile} at {company_name} ({start_date.strftime('%B %Y')} - {'Present' if is_current_job else end_date.strftime('%B %Y')})"
    add_and_style_cell(experience_text)
    add_and_style_cell(jd)  # Assuming 'jd' contains the job description

    add_and_style_cell('Education', bold=True)
    education_text = f"{degree} from {university}"
    add_and_style_cell(education_text)

    # Handle certifications and additional skills similarly
    add_and_style_cell('Certifications', bold=True)
    for cert in certifications.split('\n'):
        add_and_style_cell(cert)

    add_and_style_cell('Additional Skills', bold=True)
    for skill in additional_skills.split('\n'):
        add_and_style_cell(skill)

    # Save the document
    file_path = os.path.join(os.getcwd(), 'resume.docx')
    doc.save(file_path)
    return file_path


def main():
    st.title('ATS Friendly Resume Generator')

    with st.form("resume_form"):
        with st.expander("Personal Details"):
            name = st.text_input("Name")
            city = st.text_input("City")
            area_name = st.text_input("Area")
            zipcode = st.text_input("Zipcode")
            email = st.text_input("Email")
            phone = st.text_input("Phone")
            linkedin = st.text_input("LinkedIn")
            summary = st.text_area("Summary", placeholder="Write a Brief Summary")
        with st.expander('Technical Skills'):
            statistical_methods = st.multiselect(label='Statistical methods',
                                                 options=['Statistical Techniques', 'Descriptive Statistics',
                                                          'Inferential Statistics', 'Probability Distribution',
                                                          'Hypothesis Testing', 'Regression'])
            programming_languages = st.multiselect(label='Programming Languages', options=['Python', 'SQL'])
            data_collection = st.multiselect(label='Data Collection',
                                             options=['requests', 'bs4', 'BeautifulSoup', 'lxml', 'API', 'web scraping'])
            libraries = st.multiselect(label='Libraries',
                                       options=['Numpy', 'Pandas', 'Matplotlib', 'Seaborn', 'Plotly', 'OpenCV'])
            database_management = st.multiselect(label='Database Management', options=['MySQL', 'MS SQL Server'])
            business_intelligence = st.multiselect(label='Business Intelligence',
                                                   options=['Power BI', 'DAX', 'Power Query', 'Tableau', 'Zoho',
                                                            'Quicksight', 'Google Studio', 'Excel Reporting'])
            data_engineering = st.multiselect(label='Data Engineering',
                                              options=['ETL Tools', 'SSIS', 'Data Warehouse', 'Data Pipeline',
                                                       'Data Mining', 'Data Wrangling', 'Data Munging'])
            cloud_platforms = st.multiselect(label='Cloud Platform',
                                             options=['Microsoft Azure', 'Azure Data Factory', 'Azure Cloud Services',
                                                      'Azure SQL Database', 'AWS', 'GCP'])
            big_data = st.multiselect(label='Big Data Tools', options=['Apache Spark', 'PySpark', 'Databricks'])
            machine_learning = st.multiselect(label='Machine Learning', options=['Scikit-learn'])
            profile = st.text_input('Previous Profile')
        with st.expander('Professional Experience'):
            company_name = st.text_input('Company Name')
            # Input fields for start and end date, and a checkbox for the current job
            start_date = st.date_input("Start Date")
            end_date = st.date_input("End Date")
            is_current_job = st.checkbox("Currently Working Here")
            jd = st.text_area('Explain Job Description')
        with st.expander('Education'):
            degree = st.text_input('Education')
            university = st.text_input('University')
            certifications = st.text_area('Certifications')
            additional_skills = st.text_area('Additional Skills')

        # Other expanders for Skills, Experience, etc.
        # Make sure to collect all required information as done in the Personal Details section

        submitted = st.form_submit_button("Generate Resume")

    if submitted:
        # Call the function to generate resume
        # Make sure to pass all collected information to the function
        file_path = generate_resume(name, city, area_name, zipcode, email, phone, linkedin,
                                    summary, programming_languages, libraries, business_intelligence,
                                    data_engineering, big_data, profile, company_name,
                                    start_date, end_date, is_current_job, jd, degree, university,
                                    certifications, additional_skills, statistical_methods,
                                    data_collection, database_management, cloud_platforms,
                                    machine_learning)

        st.session_state['resume_path'] = file_path  # Store the path in session state

    # Check if the resume has been generated and offer a download button
    if 'resume_path' in st.session_state:
        with open(st.session_state['resume_path'], 'rb') as file:
            st.download_button(label="Download Resume", data=file, file_name="resume.docx",
                               mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


if __name__ == '__main__':
    main()



















#-------------------------------------------------------------------------------------------------------------------------------------
# import streamlit as st
# from docx import Document
# from docx.shared import Mm
# from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
# import os


# def generate_resume(name, city, area_name, zipcode, email, phone, linkedin,
#                     summary, programming_languages, libraries, business_intelligence,
#                     data_engineering, big_data, profile, company_name,
#                     start_date, end_date, is_current_job, jd, degree, university, certifications,
#                     additional_skills, statistical_methods, data_collection,
#                     database_management, cloud_platforms, machine_learning):
#     doc = Document()

#     # Set narrow margins
#     sections = doc.sections
#     for section in sections:
#         section.top_margin = Mm(10)
#         section.bottom_margin = Mm(10)
#         section.left_margin = Mm(10)
#         section.right_margin = Mm(10)

#     # Add a table for the entire resume
#     table = doc.add_table(rows=19, cols=1)  # Adjusted row count for the date fields
#     table.style = 'Table Grid'

#     # Name
#     cell = table.cell(0, 0)
#     cell.text = name
#     cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
#     cell.paragraphs[0].runs[0].font.bold = True

#     # Contact information
#     cell = table.cell(1, 0)
#     cell.text = 'Data Analyst'
#     cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
#     cell = table.cell(2, 0)
#     cell.text = f"{city},{area_name},{zipcode} | {email} | {phone} | {linkedin}"
#     cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

#     # Summary
#     cell = table.cell(3, 0)
#     cell.text = 'Summary'
#     cell.paragraphs[0].runs[0].font.bold = True
#     cell = table.cell(4, 0)
#     cell.text = summary

#     # Skills
#     cell = table.cell(5, 0)
#     cell.text = 'Technical Skills'
#     cell.paragraphs[0].runs[0].font.bold = True
#     cell = table.cell(6, 0)
#     cell.text = f"• Statistical Methods:{','.join(statistical_methods)}\n• Programming Languages:{','.join(programming_languages)}\n• Data Collection:{','.join(data_collection)}\n• Libraries:{','.join(libraries)}\n• Database Management:{','.join(database_management)}\n• Business Intelligence:{','.join(business_intelligence)}\n• Data Engineering:{','.join(cloud_platforms)}\n• Data Engineering:{','.join(data_engineering)}\n• Big Data:{','.join(big_data)}\n• Machine Learning:{','.join(machine_learning)}"

#     # Company and Profile Info
#     cell = table.cell(7, 0)
#     cell.text = 'Professional Experience'
#     cell.paragraphs[0].runs[0].font.bold = True
#     # Adjusting the Professional Experience section to include dates
#     cell = table.cell(8, 0)
#     end_date_text = "Present" if is_current_job else end_date.strftime('%B %Y')  # Formatting date
#     cell.text = f"{profile}\n{company_name}\n{start_date.strftime('%B %Y')} - {end_date_text}"
#     cell.paragraphs[0].runs[0].font.bold = True

#     # Experience Description
#     cell = table.cell(9, 0)
#     # cell.paragraphs[0].runs[0].font.bold = True
#     cell.text = '\n'.join(['• ' + j for j in jd.split("\n")])

#     # Continue with other sections like Education, Certifications, etc...
#     # Education
#     cell = table.cell(10, 0)
#     cell.text = 'Education'
#     cell.paragraphs[0].runs[0].font.bold = True
#     cell = table.cell(11, 0)
#     cell.text = f'{degree}\n{university}'

#     # Certifications
#     cell = table.cell(12, 0)
#     cell.text = 'Certifications'
#     cell.paragraphs[0].runs[0].font.bold = True
#     cell = table.cell(13, 0)
#     cell.text = '\n'.join(['• ' + c for c in certifications.split("\n")])

#     # Additional Skills
#     cell = table.cell(14, 0)
#     cell.text = 'Additional Skills'
#     cell.paragraphs[0].runs[0].font.bold = True
#     cell = table.cell(15, 0)
#     cell.text = '\n'.join(['• ' + ad for ad in additional_skills.split("\n")])

#     # Specify the absolute path to save the document
#     file_path = os.path.join(os.getcwd(), 'resume.docx')
#     doc.save(file_path)
#     return file_path


# def main():
#     st.title('ATS Friendly Resume Generator')

#     # UI for input fields
#     name = st.text_input('Name')
#     city = st.text_input('City')
#     area_name = st.text_input('Area')
#     zipcode = st.text_input('Zipcode')
#     email = st.text_input('Email')
#     phone = st.text_input('Phone')
#     linkedin = st.text_input('LinkedIn')
#     summary = st.text_area('Summary', placeholder='Write a Brief Summary')
#     statistical_methods = st.multiselect(label='Statistical methods',options=['Statistical Techniques','Descriptive Statistics','Inferential Statistics','Probability Distribution','Hypothesis Testing','Regression'])
#     programming_languages = st.multiselect(label='Programming Languages',options=['Python','SQL'])
#     data_collection = st.multiselect(label='Data Collection',options=['requests','bs4','BeautifulSoup','lxml','API','web scraping'])
#     libraries = st.multiselect(label='Libraries',options=['Numpy','Pandas','Matplotlib','Seaborn','Plotly','OpenCV'])
#     database_management = st.multiselect(label='Database Management',options=['MySQL','MS SQL Server'])
#     business_intelligence = st.multiselect(label='Business Intelligence',options=['Power BI','DAX','Power Query','Tableau','Zoho','Quicksight','Google Studio','Excel Reporting'])
#     data_engineering = st.multiselect(label='Data Engineering',options=['ETL Tools', 'SSIS', 'Data Warehouse', 'Data Pipeline','Data Mining','Data Wrangling','Data Munging'])
#     cloud_platforms = st.multiselect(label='Cloud Platform',options=['Microsoft Azure','Azure Data Factory','Azure Cloud Services','Azure SQL Database','AWS','GCP'])
#     big_data = st.multiselect(label='Big Data Tools',options=['Apache Spark','PySpark','Databricks'])
#     machine_learning = st.multiselect(label='Machine Learning',options=['Scikit-learn'])
#     profile = st.text_input('Previous Profile')
#     company_name = st.text_input('Company Name')
#     # Input fields for start and end date, and a checkbox for the current job
#     start_date = st.date_input("Start Date")
#     end_date = st.date_input("End Date")
#     is_current_job = st.checkbox("Currently Working Here")
#     jd =st.text_area('Explain Job Description')
#     degree = st.text_input('Education')
#     university = st.text_input('University')
#     certifications = st.text_area('Certifications')
#     additional_skills = st.text_area('Additional Skills')




#     # The rest of your multiselect and input fields for skills, experience, etc.

#     if st.button('Generate Resume'):
#         # Check if all necessary fields are filled
#         if name and city and area_name and zipcode and email and phone and linkedin and summary:
#             file_path = generate_resume(name, city, area_name, zipcode, email, phone, linkedin,
#                                         summary, programming_languages, libraries, business_intelligence,
#                                         data_engineering, big_data, profile, company_name, start_date,
#                                         end_date, is_current_job, jd, degree, university, certifications,
#                                         additional_skills, statistical_methods, data_collection,
#                                         database_management, cloud_platforms, machine_learning)

#             st.success('Resume generated successfully!')
#             with open(file_path, 'rb') as file:
#                 st.download_button(label="Download your resume", data=file, file_name='resume.docx',
#                                    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


# if __name__ == '__main__':
#     main()
