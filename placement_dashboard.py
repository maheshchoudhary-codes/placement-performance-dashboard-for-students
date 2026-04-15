import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import base64, os 
import webbrowser
from io import BytesIO
import warnings
warnings.filterwarnings('ignore')


# load or create data
def load_data():
    if not os.path.exists("placement_data.xlsx"):
        print("file not found, making sample data...")
        make_data()
    df = pd.read_excel("placement_data.xlsx")
    return df


def make_data():
    np.random.seed(42)

    depts = ['Data Science', 'Computer Science', 'IT', 'Electronics', 'Mechanical', 'Civil']

    it_companies = ['TCS', 'Infosys', 'Wipro', 'Accenture', 'Cognizant',
                    'HCL', 'Tech Mahindra', 'Capgemini', 'IBM', 'Oracle',
                    'Persistent', 'Mphasis', 'Hexaware', 'Zensar']
    core_companies = ['Tata Motors', 'L&T', 'BHEL', 'Siemens', 'Bosch',
                      'Cummins', 'Thermax', 'Kirloskar']

    it_roles = ['Software Engineer', 'Data Analyst', 'System Engineer',
                'Jr Developer', 'QA Engineer', 'Python Developer', 'ML Engineer']
    core_roles = ['Graduate Engineer Trainee', 'Design Engineer',
                  'Site Engineer', 'Production Engineer']

    all_skills = ['Python', 'Java', 'SQL', 'Machine Learning', 'Power BI',
                  'Excel', 'C++', 'JavaScript', 'Django', 'Data Analysis',
                  'Tableau', 'R', 'Deep Learning', 'AutoCAD', 'MATLAB',
                  'Embedded C', 'IoT', 'AWS', 'Linux', 'React']

    fnames = ['Rahul', 'Priya', 'Amit', 'Sneha', 'Rohan', 'Pooja', 'Akash',
              'Neha', 'Vishal', 'Ankita', 'Sagar', 'Riya', 'Nikhil', 'Divya',
              'Pratik', 'Swati', 'Harsh', 'Komal', 'Vaibhav', 'Shruti',
              'Omkar', 'Pallavi', 'Tejas', 'Rutuja', 'Gaurav', 'Shubham', 'Kunal']
    lnames = ['Patil', 'Shinde', 'Jadhav', 'Deshmukh', 'More', 'Kulkarni',
              'Gaikwad', 'Pawar', 'Bhosale', 'Wagh', 'Mane', 'Kale',
              'Deshpande', 'Sawant', 'Yadav', 'Choudhary', 'Nikam', 'Kamble',
              'Thorat', 'Bansode', 'Suryawanshi', 'Waghmare']

    dept_count = {'Data Science':40, 'Computer Science':45, 'IT':38,
                  'Electronics':30, 'Mechanical':35, 'Civil':25}

    rows = []
    roll = 1
    for dept, cnt in dept_count.items():
        for i in range(cnt):
            cgpa  = round(np.random.uniform(5.2, 9.8), 1)
            skill = round(np.random.uniform(4.5, 9.5), 1)

            # simple placement probability
            prob = 0.5
            if cgpa >= 7.5:  prob += 0.25
            if skill >= 7.5: prob += 0.15
            if cgpa < 6 and skill < 6: prob = 0.2

            placed = np.random.random() < prob
            is_it  = dept in ['Data Science', 'Computer Science', 'IT']

            if placed:
                comp = np.random.choice(it_companies if is_it else core_companies)
                role = np.random.choice(it_roles if is_it else core_roles)
                pkg  = round(np.random.uniform(3.5, 14.0), 1)
                if comp in ['Oracle', 'IBM']:
                    pkg = round(np.random.uniform(8.0, 16.0), 1)
                status = 'Placed'
            else:
                comp, role, pkg, status = '-', '-', 0, 'Not Placed'

            s1, s2, s3 = np.random.choice(all_skills, 3, replace=False)
            yr = np.random.choice([2021, 2022, 2023, 2024, 2025])

            rows.append({
                'Roll_No':          f'DC{str(roll).zfill(3)}',
                'Student_Name':     np.random.choice(fnames) + ' ' + np.random.choice(lnames),
                'Department':       dept,
                'CGPA':             cgpa,
                'Skill_Score':      skill,
                'Skill_1':          s1,
                'Skill_2':          s2,
                'Skill_3':          s3,
                'Company':          comp,
                'Job_Role':         role,
                'Package_LPA':      pkg,
                'Placement_Status': status,
                'Year':             yr
            })
            roll += 1

    pd.DataFrame(rows).to_excel("placement_data.xlsx", index=False)
    print(f"created {len(rows)} student records")


# img to base64 so we can embed in html
def img_b64():
    buf = BytesIO()
    plt.savefig(buf, format='png', dpi=120, bbox_inches='tight')
    buf.seek(0)
    data = base64.b64encode(buf.read()).decode()
    plt.close()
    return data


def run_dashboard():
    df = load_data()

    # basic cleaning
    df['Package_LPA']  = pd.to_numeric(df['Package_LPA'],  errors='coerce').fillna(0)
    df['CGPA']         = pd.to_numeric(df['CGPA'],         errors='coerce').fillna(0)
    df['Skill_Score']  = pd.to_numeric(df['Skill_Score'],  errors='coerce').fillna(0)
    df['Placement_Status'] = df['Placement_Status'].str.strip()

    df['is_placed']      = df['Placement_Status'].str.lower() == 'placed'
    df['skills_case']    = (df['CGPA'] < 7.0) & (df['Skill_Score'] >= 7.5)

    placed     = df[df['is_placed']]
    not_placed = df[~df['is_placed']]
    skill_df   = df[df['skills_case'] & df['is_placed']]

    total    = len(df)
    n_placed = len(placed)
    rate     = round(n_placed / total * 100, 1)
    avg_pkg  = round(placed['Package_LPA'].mean(), 2)
    max_pkg  = placed['Package_LPA'].max()

    print(f"Total: {total}  |  Placed: {n_placed}  |  Rate: {rate}%  |  Avg Pkg: {avg_pkg} LPA")

    charts = {}

    # -- chart 1: placement status pie --
    fig, ax = plt.subplots(figsize=(6, 5))
    ax.pie([n_placed, total - n_placed],
           labels=['Placed', 'Not Placed'],
           autopct='%1.1f%%',
           colors=['#1F3864', '#E74C3C'],
           startangle=90,
           wedgeprops=dict(width=0.5, edgecolor='white'))
    ax.text(0, 0, f"{n_placed}/{total}", ha='center', va='center',
            fontsize=11, fontweight='bold', color='#1F3864')
    ax.set_title('Placement Status', fontsize=13, fontweight='bold')
    charts['c1'] = img_b64()

    # -- chart 2: dept wise bar --
    dept_g = df.groupby('Department')['is_placed'].agg(['sum','count'])
    dept_g['rate'] = (dept_g['sum'] / dept_g['count'] * 100).round(1)
    dept_g = dept_g.sort_values('rate')

    fig, ax = plt.subplots(figsize=(8, 5))
    ax.barh(dept_g.index, dept_g['rate'], color='#2E75B6', edgecolor='white')
    for i, v in enumerate(dept_g['rate']):
        ax.text(v + 0.5, i, f'{v}%', va='center', fontsize=9, color='#1F3864')
    ax.set_xlabel('Placement Rate (%)')
    ax.set_title('Dept-wise Placement Rate', fontsize=13, fontweight='bold')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    charts['c2'] = img_b64()

    # -- chart 3: cgpa vs skill scatter --
    fig, ax = plt.subplots(figsize=(7, 5))
    ax.scatter(not_placed['CGPA'], not_placed['Skill_Score'],
               c='#E74C3C', s=55, alpha=0.6, label='Not Placed', marker='x')
    ax.scatter(placed[~placed['skills_case']]['CGPA'],
               placed[~placed['skills_case']]['Skill_Score'],
               c='#1F3864', s=55, alpha=0.7, label='Placed')
    ax.scatter(skill_df['CGPA'], skill_df['Skill_Score'],
               c='#FFC000', s=90, alpha=0.9, label='Skills-based', marker='*')
    ax.axvline(7.0, color='gray', linestyle='--', alpha=0.5)
    ax.axhline(7.5, color='gray', linestyle='--', alpha=0.5)
    ax.set_xlabel('CGPA')
    ax.set_ylabel('Skill Score')
    ax.set_title('CGPA vs Skill Score', fontsize=13, fontweight='bold')
    ax.legend(fontsize=9)
    ax.grid(alpha=0.2)
    charts['c3'] = img_b64()

    # -- chart 4: top companies --
    top_co = placed[placed['Company'] != '-']['Company'].value_counts().head(10)

    fig, ax = plt.subplots(figsize=(8, 5))
    ax.barh(top_co.index[::-1], top_co.values[::-1], color='#1F3864', edgecolor='white')
    ax.set_xlabel('Students Placed')
    ax.set_title('Top Recruiting Companies', fontsize=13, fontweight='bold')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    charts['c4'] = img_b64()

    # -- chart 5: salary distribution --
    pkgs = placed[placed['Package_LPA'] > 0]['Package_LPA']

    fig, ax = plt.subplots(figsize=(7, 5))
    ax.hist(pkgs, bins=10, color='#2E75B6', edgecolor='white')
    ax.axvline(pkgs.mean(), color='#FFC000', linestyle='--', linewidth=2,
               label=f'Avg: {pkgs.mean():.2f} LPA')
    ax.set_xlabel('Package (LPA)')
    ax.set_ylabel('No. of Students')
    ax.set_title('Salary Distribution', fontsize=13, fontweight='bold')
    ax.legend()
    ax.grid(axis='y', alpha=0.3)
    charts['c5'] = img_b64()

    # dept summary table rows
    dept_tbl = df.groupby('Department').agg(
        Total   =('Roll_No',       'count'),
        Placed  =('is_placed',     'sum'),
        AvgCGPA =('CGPA',          'mean'),
        AvgSkill=('Skill_Score',   'mean'),
        AvgPkg  =('Package_LPA',   'mean')
    ).round(2)
    dept_tbl['Rate'] = ((dept_tbl['Placed'] / dept_tbl['Total']) * 100).round(1)

    dept_rows_html = ''
    for dept, r in dept_tbl.iterrows():
        dept_rows_html += f"<tr><td>{dept}</td><td>{int(r.Total)}</td><td>{int(r.Placed)}</td><td>{r.Rate}%</td><td>{r.AvgCGPA}</td><td>{r.AvgPkg}</td></tr>"

    skill_rows_html = ''
    for _, r in skill_df.head(8).iterrows():
        skill_rows_html += f"<tr><td>{r.Student_Name}</td><td>{r.Department}</td><td>{r.CGPA}</td><td>{r.Skill_Score}</td><td>{r.Company}</td><td>{r.Package_LPA} LPA</td></tr>"

    # build html
    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Placement Dashboard - Deogiri College</title>
<style>
  body {{ font-family: Arial, sans-serif; margin:0; background:#f0f4f8; color:#222; font-size:14px; }}
  .header {{ background:#1F3864; color:white; padding:24px 36px; border-bottom:4px solid #FFC000; }}
  .header h1 {{ font-size:22px; margin:0; }}
  .header p  {{ font-size:12px; margin-top:5px; opacity:0.8; }}
  .kpis {{
    display:grid; grid-template-columns:repeat(5,1fr);
    gap:14px; padding:20px 36px; background:white; border-bottom:1px solid #ddd;
  }}
  .kpi {{ background:#1F3864; border-radius:8px; padding:16px; text-align:center; }}
  .kpi .val {{ font-size:26px; font-weight:700; color:#FFC000; }}
  .kpi .lbl {{ font-size:10px; color:#aac; margin-top:3px; text-transform:uppercase; }}
  .section  {{ padding:24px 36px; }}
  .section h2 {{ font-size:15px; font-weight:700; color:#1F3864;
                 border-left:4px solid #FFC000; padding-left:10px; margin-bottom:16px; }}
  .grid2 {{ display:grid; grid-template-columns:1fr 1fr; gap:16px; }}
  .cbox  {{ background:white; border-radius:8px; padding:14px; border:1px solid #ddd; }}
  .cbox img {{ width:100%; }}
  table  {{ width:100%; border-collapse:collapse; background:white;
            border-radius:8px; overflow:hidden; }}
  th {{ background:#1F3864; color:white; padding:10px 12px; font-size:12px; text-align:left; }}
  td {{ padding:9px 12px; font-size:13px; border-bottom:1px solid #eee; }}
  tr:nth-child(even) td {{ background:#f7f9fc; }}
  .footer {{ background:#1F3864; color:#aac; text-align:center; padding:14px; font-size:12px; }}
</style>
</head>
<body>

<div class="header">
  <h1>Placement Performance Dashboard</h1>
  <p>Deogiri College, Chhatrapati Sambhajinagar &nbsp;|&nbsp; B.Sc Data Science (SY) &nbsp;|&nbsp; AY 2025-26 &nbsp;|&nbsp; Guide: Mr. Dilip B. Dabhade</p>
</div>

<div class="kpis">
  <div class="kpi"><div class="val">{total}</div><div class="lbl">Total Students</div></div>
  <div class="kpi"><div class="val">{n_placed}</div><div class="lbl">Placed</div></div>
  <div class="kpi"><div class="val">{rate}%</div><div class="lbl">Placement Rate</div></div>
  <div class="kpi"><div class="val">Rs.{avg_pkg}</div><div class="lbl">Avg Package</div></div>
  <div class="kpi"><div class="val">{len(skill_df)}</div><div class="lbl">Skills-based</div></div>
</div>

<div class="section">
  <h2>Visualizations</h2>
  <div class="grid2">
    <div class="cbox"><img src="data:image/png;base64,{charts['c1']}"/></div>
    <div class="cbox"><img src="data:image/png;base64,{charts['c2']}"/></div>
    <div class="cbox"><img src="data:image/png;base64,{charts['c3']}"/></div>
    <div class="cbox"><img src="data:image/png;base64,{charts['c4']}"/></div>
  </div>
  <br>
  <div class="cbox" style="max-width:600px">
    <img src="data:image/png;base64,{charts['c5']}"/>
  </div>
</div>

<div class="section">
  <h2>Department Summary</h2>
  <table>
    <thead><tr><th>Department</th><th>Total</th><th>Placed</th><th>Rate</th><th>Avg CGPA</th><th>Avg Pkg</th></tr></thead>
    <tbody>{dept_rows_html}</tbody>
  </table>
</div>

<div class="section">
  <h2>Skills-based Placements (CGPA &lt; 7.0 and Skill ≥ 7.5)</h2>
  <table>
    <thead><tr><th>Name</th><th>Dept</th><th>CGPA</th><th>Skill Score</th><th>Company</th><th>Package</th></tr></thead>
    <tbody>{skill_rows_html}</tbody>
  </table>
</div>

<div class="footer">
  Placement Dashboard &nbsp;|&nbsp; Deogiri College &nbsp;|&nbsp; Data Science Dept &nbsp;|&nbsp; 2025-26
</div>

</body>
</html>"""

    with open("placement_dashboard.html", "w", encoding="utf-8") as f:
        f.write(html)

    print("Done! Open placement_dashboard.html in browser.")


run_dashboard()
webbrowser.open(os.path.abspath("placement_dashboard.html"))

