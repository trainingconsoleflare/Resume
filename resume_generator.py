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
    skill_lines = [f"Programming Languages: {', '.join(programming_languages)}",
                   f"Libraries: {', '.join(libraries)}",
                   f"Business Intelligence: {', '.join(business_intelligence)}",
                   f"Data Engineering: {', '.join(data_engineering)}",
                   f"Big Data Tools: {', '.join(big_data)}",
                   f"Statistical Methods: {', '.join(statistical_methods)}",
                   f"Data Collection Techniques: {', '.join(data_collection)}",
                   f"Database Management Systems: {', '.join(database_management)}",
                   f"Cloud Platforms: {', '.join(cloud_platforms)}",
                   f"Machine Learning Libraries: {', '.join(machine_learning)}"]
    # Filter out empty lines
    skill_lines = [line for line in skill_lines if not line.endswith(': ')]
    skills_text = '\n'.join(skill_lines)
    add_and_style_cell(skills_text)

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











