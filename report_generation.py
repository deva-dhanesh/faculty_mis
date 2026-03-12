"""
Faculty Report Generation Module
Generates visual representations and text interpretations of faculty academic data.
"""

import json
from datetime import datetime
import plotly.graph_objects as go
import plotly.express as px
from models import (
    User, FacultyPublication, FacultyProject, FacultyBookChapter, FacultyPatent,
    FacultyAward, ConferenceParticipated, ConferenceOrganised, 
    FDPParticipated, FDPOrganised, CourseAttended, CourseOffered, FacultyGuestLecture,
    FacultyFellowship
)


def get_publications_data(user_id):
    """Fetch publications data"""
    pubs = FacultyPublication.query.filter_by(user_id=user_id).all()
    return {
        'count': len(pubs),
        'by_type': {},
        'by_quartile': {},
        'by_indexing': {},
        'by_year': {},
        'items': pubs
    }


def get_projects_data(user_id):
    """Fetch projects data"""
    projs = FacultyProject.query.filter_by(user_id=user_id).all()
    status_count = {}
    type_count = {}
    
    for proj in projs:
        status_count[proj.status] = status_count.get(proj.status, 0) + 1
        type_count[proj.project_type] = type_count.get(proj.project_type, 0) + 1
    
    return {
        'count': len(projs),
        'by_status': status_count,
        'by_type': type_count,
        'items': projs
    }


def get_books_data(user_id):
    """Fetch books and chapters data"""
    books = FacultyBookChapter.query.filter_by(user_id=user_id).all()
    return {
        'count': len(books),
        'items': books
    }


def get_patents_data(user_id):
    """Fetch patents data"""
    patents = FacultyPatent.query.filter_by(user_id=user_id).all()
    status_count = {}
    
    for patent in patents:
        status_count[patent.status] = status_count.get(patent.status, 0) + 1
    
    return {
        'count': len(patents),
        'by_status': status_count,
        'items': patents
    }


def get_awards_data(user_id):
    """Fetch awards data"""
    awards = FacultyAward.query.filter_by(user_id=user_id).all()
    return {
        'count': len(awards),
        'items': awards
    }


def get_conferences_data(user_id):
    """Fetch conferences data"""
    conf_part = ConferenceParticipated.query.filter_by(user_id=user_id).count()
    conf_org = ConferenceOrganised.query.filter_by(user_id=user_id).count()
    
    return {
        'participated': conf_part,
        'organised': conf_org,
        'total': conf_part + conf_org
    }


def get_fdp_data(user_id):
    """Fetch FDP data"""
    fdp_part = FDPParticipated.query.filter_by(user_id=user_id).count()
    fdp_org = FDPOrganised.query.filter_by(user_id=user_id).count()
    
    return {
        'participated': fdp_part,
        'organised': fdp_org,
        'total': fdp_part + fdp_org
    }


def get_courses_data(user_id):
    """Fetch courses data"""
    courses_att = CourseAttended.query.filter_by(user_id=user_id).count()
    courses_off = CourseOffered.query.filter_by(user_id=user_id).count()
    
    return {
        'attended': courses_att,
        'offered': courses_off,
        'total': courses_att + courses_off
    }


def get_guest_lectures_data(user_id):
    """Fetch guest lectures data"""
    lectures = FacultyGuestLecture.query.filter_by(user_id=user_id).count()
    return {'count': lectures}


def get_fellowships_data(user_id):
    """Fetch fellowships data"""
    fellowships = FacultyFellowship.query.filter_by(user_id=user_id).count()
    return {'count': fellowships}


def generate_publication_chart(user_id):
    """Generate publication distribution chart"""
    pubs = FacultyPublication.query.filter_by(user_id=user_id).all()
    
    if not pubs:
        return None
    
    # Count by type
    type_counts = {}
    for pub in pubs:
        pub_type = pub.publication_type or "Unspecified"
        type_counts[pub_type] = type_counts.get(pub_type, 0) + 1
    
    fig = go.Figure(data=[
        go.Bar(x=list(type_counts.keys()), y=list(type_counts.values()),
               marker=dict(color='#2563eb'))
    ])
    fig.update_layout(
        title="Publications by Type",
        xaxis_title="Publication Type",
        yaxis_title="Count",
        height=400,
        hovermode='x unified',
        margin=dict(l=40, r=20, t=40, b=40),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)'
    )
    return fig.to_html(div_id="pub_chart", include_plotlyjs='cdn')


def generate_project_chart(user_id):
    """Generate project status chart"""
    projs = FacultyProject.query.filter_by(user_id=user_id).all()
    
    if not projs:
        return None
    
    # Count by status
    status_counts = {}
    for proj in projs:
        status = proj.status or "Not Specified"
        status_counts[status] = status_counts.get(status, 0) + 1
    
    colors = {'Ongoing': '#16a34a', 'Completed': '#2563eb', 'Not Specified': '#94a3b8'}
    
    fig = go.Figure(data=[
        go.Pie(labels=list(status_counts.keys()), values=list(status_counts.values()),
               marker=dict(colors=[colors.get(k, '#64748b') for k in status_counts.keys()]))
    ])
    fig.update_layout(
        title="Project Status Distribution",
        height=400,
        margin=dict(l=40, r=20, t=40, b=40)
    )
    return fig.to_html(div_id="proj_chart", include_plotlyjs=False)


def generate_activity_chart(user_id):
    """Generate activity comparison chart"""
    data_dict = {
        'Publications': FacultyPublication.query.filter_by(user_id=user_id).count(),
        'Projects': FacultyProject.query.filter_by(user_id=user_id).count(),
        'Books': FacultyBookChapter.query.filter_by(user_id=user_id).count(),
        'Patents': FacultyPatent.query.filter_by(user_id=user_id).count(),
        'Awards': FacultyAward.query.filter_by(user_id=user_id).count(),
        'Conferences': (ConferenceParticipated.query.filter_by(user_id=user_id).count() + 
                       ConferenceOrganised.query.filter_by(user_id=user_id).count()),
        'FDP': (FDPParticipated.query.filter_by(user_id=user_id).count() + 
               FDPOrganised.query.filter_by(user_id=user_id).count()),
    }
    
    # Filter out zero values
    data_dict = {k: v for k, v in data_dict.items() if v > 0}
    
    if not data_dict:
        return None
    
    fig = go.Figure(data=[
        go.Bar(x=list(data_dict.keys()), y=list(data_dict.values()),
               marker=dict(color='#2563eb', opacity=0.8))
    ])
    fig.update_layout(
        title="Overall Academic Activity Summary",
        xaxis_title="Activity Type",
        yaxis_title="Count",
        height=400,
        hovermode='x unified',
        margin=dict(l=40, r=20, t=40, b=40),
        xaxis_tickangle=-45
    )
    return fig.to_html(div_id="activity_chart", include_plotlyjs=False)


def generate_interpretation(user_id, features):
    """Generate text interpretation of report data"""
    
    interpretation = "<div style='color:#475569;line-height:1.8;'>"
    
    if 'publications' in features:
        pubs_count = FacultyPublication.query.filter_by(user_id=user_id).count()
        if pubs_count > 0:
            pubs = FacultyPublication.query.filter_by(user_id=user_id).all()
            q1_count = sum(1 for p in pubs if p.journal_quartile == 'Q1')
            q2_count = sum(1 for p in pubs if p.journal_quartile == 'Q2')
            
            interpretation += f"""
            <p><strong>Publication Profile:</strong> You have published <strong>{pubs_count} research articles</strong> 
            across various journals. """
            
            if q1_count > 0:
                interpretation += f"Of these, {q1_count} are in Q1-ranked journals. "
            if q2_count > 0:
                interpretation += f"{q2_count} are in Q2-ranked journals. "
            
            interpretation += "This demonstrates a strong research output with focus on high-quality publications.</p>"
    
    if 'projects' in features:
        proj_count = FacultyProject.query.filter_by(user_id=user_id).count()
        if proj_count > 0:
            ongoing = FacultyProject.query.filter_by(user_id=user_id, status='Ongoing').count()
            completed = FacultyProject.query.filter_by(user_id=user_id, status='Completed').count()
            
            interpretation += f"""
            <p><strong>Research Projects:</strong> You are managing <strong>{proj_count} research projects</strong>. """
            
            if ongoing > 0:
                interpretation += f"Currently, {ongoing} project(s) are ongoing "
            if completed > 0:
                interpretation += f"and {completed} have been completed, "
            
            interpretation += "reflecting consistent research activity and successful project completion.</p>"
    
    if 'books' in features:
        books_count = FacultyBookChapter.query.filter_by(user_id=user_id).count()
        if books_count > 0:
            interpretation += f"""
            <p><strong>Books & Book Chapters:</strong> You have contributed <strong>{books_count} book chapters</strong> 
            to scholarly publications, demonstrating expertise in presenting research in comprehensive formats.</p>"""
    
    if 'patents' in features:
        patents_count = FacultyPatent.query.filter_by(user_id=user_id).count()
        if patents_count > 0:
            interpretation += f"""
            <p><strong>Patents & Innovations:</strong> You hold <strong>{patents_count} patent(s)</strong>, 
            indicating contribution to innovation and technology development.</p>"""
    
    if 'awards' in features:
        awards_count = FacultyAward.query.filter_by(user_id=user_id).count()
        if awards_count > 0:
            interpretation += f"""
            <p><strong>Awards & Recognition:</strong> You have received <strong>{awards_count} award(s)</strong> 
            and recognitions, reflecting your contributions to academia and research.</p>"""
    
    if 'conferences' in features:
        conf_count = ConferenceParticipated.query.filter_by(user_id=user_id).count()
        conf_org_count = ConferenceOrganised.query.filter_by(user_id=user_id).count()
        
        if conf_count > 0 or conf_org_count > 0:
            interpretation += f"""
            <p><strong>Conference Engagement:</strong> You have participated in <strong>{conf_count} conference(s)</strong> 
            and organized {conf_org_count} conference(s), demonstrating active involvement in the academic community.</p>"""
    
    if 'fdp' in features:
        fdp_count = FDPParticipated.query.filter_by(user_id=user_id).count()
        fdp_org_count = FDPOrganised.query.filter_by(user_id=user_id).count()
        
        if fdp_count > 0 or fdp_org_count > 0:
            interpretation += f"""
            <p><strong>Faculty Development:</strong> You have attended <strong>{fdp_count} FDP program(s)</strong> 
            and organized {fdp_org_count} FDP(s), showing commitment to continuous professional development.</p>"""
    
    interpretation += """
            <p style='margin-top:16px;padding-top:16px;border-top:1px solid #e2e8f0;color:#64748b;font-size:13px;'>
            <strong>Overall Assessment:</strong> Your profile demonstrates consistent academic engagement across multiple dimensions. 
            The diversity of activities—from publications and projects to conferences and professional development—
            indicates a well-rounded contribution to academic and research excellence.
            </p>
            </div>
    """
    
    return interpretation


def compile_summary(user_id, features):
    """Compile summary statistics"""
    summary = {}
    
    if 'publications' in features:
        summary['total_publications'] = FacultyPublication.query.filter_by(user_id=user_id).count()
    
    if 'projects' in features:
        summary['total_projects'] = FacultyProject.query.filter_by(user_id=user_id).count()
    
    if 'books' in features:
        summary['total_books'] = FacultyBookChapter.query.filter_by(user_id=user_id).count()
    
    if 'patents' in features:
        summary['total_patents'] = FacultyPatent.query.filter_by(user_id=user_id).count()
    
    if 'awards' in features:
        summary['total_awards'] = FacultyAward.query.filter_by(user_id=user_id).count()
    
    if 'conferences' in features:
        conf_part = ConferenceParticipated.query.filter_by(user_id=user_id).count()
        conf_org = ConferenceOrganised.query.filter_by(user_id=user_id).count()
        summary['total_conferences'] = conf_part + conf_org
    
    if 'fdp' in features:
        fdp_part = FDPParticipated.query.filter_by(user_id=user_id).count()
        fdp_org = FDPOrganised.query.filter_by(user_id=user_id).count()
        summary['total_fdp'] = fdp_part + fdp_org
    
    return summary


def generate_charts(user_id, features):
    """Generate all relevant charts"""
    charts = []
    
    # Activity overview chart (always include if any feature is selected)
    activity_chart = generate_activity_chart(user_id)
    if activity_chart:
        charts.append(activity_chart)
    
    if 'publications' in features:
        pub_chart = generate_publication_chart(user_id)
        if pub_chart:
            charts.append(pub_chart)
    
    if 'projects' in features:
        proj_chart = generate_project_chart(user_id)
        if proj_chart:
            charts.append(proj_chart)
    
    return charts


def generate_detailed_stats(user_id, features):
    """Generate detailed statistics for each category"""
    detailed = {}
    
    if 'publications' in features:
        pubs = FacultyPublication.query.filter_by(user_id=user_id).all()
        if pubs:
            detailed['Publications'] = {
                'Total Count': len(pubs),
                'Q1 Journals': sum(1 for p in pubs if p.journal_quartile == 'Q1'),
                'Q2 Journals': sum(1 for p in pubs if p.journal_quartile == 'Q2'),
                'Q3 Journals': sum(1 for p in pubs if p.journal_quartile == 'Q3'),
                'Q4 Journals': sum(1 for p in pubs if p.journal_quartile == 'Q4'),
            }
    
    if 'projects' in features:
        projs = FacultyProject.query.filter_by(user_id=user_id).all()
        if projs:
            detailed['Projects'] = {
                'Total Count': len(projs),
                'Ongoing': sum(1 for p in projs if p.status == 'Ongoing'),
                'Completed': sum(1 for p in projs if p.status == 'Completed'),
            }
    
    if 'conferences' in features:
        detailed['Conferences'] = {
            'Participated': ConferenceParticipated.query.filter_by(user_id=user_id).count(),
            'Organized': ConferenceOrganised.query.filter_by(user_id=user_id).count(),
        }
    
    if 'fdp' in features:
        detailed['FDP Programs'] = {
            'Participated': FDPParticipated.query.filter_by(user_id=user_id).count(),
            'Organized': FDPOrganised.query.filter_by(user_id=user_id).count(),
        }
    
    if 'courses' in features:
        detailed['Courses'] = {
            'Attended': CourseAttended.query.filter_by(user_id=user_id).count(),
            'Offered': CourseOffered.query.filter_by(user_id=user_id).count(),
        }
    
    return detailed


def generate_pdf_report(user, features, summary, interpretation, detailed_stats, current_date):
    """Generate PDF report and return file path"""
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
    import os
    import html
    from io import StringIO
    
    # Create temp directory if not exists
    temp_dir = "temp_reports"
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)
    
    pdf_path = os.path.join(temp_dir, f"report_{user.id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
    
    # Create PDF
    doc = SimpleDocTemplate(pdf_path, pagesize=letter, topMargin=0.5*inch, bottomMargin=0.5*inch)
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        textColor=colors.HexColor('#1e293b'),
        spaceAfter=6,
        alignment=TA_CENTER
    )
    
    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=colors.HexColor('#0f172a'),
        spaceAfter=12,
        spaceBefore=12,
        borderColor=colors.HexColor('#2563eb'),
        borderWidth=2,
        borderPadding=6
    )
    
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.HexColor('#475569'),
        alignment=TA_JUSTIFY,
        spaceAfter=8
    )
    
    # Build elements
    elements = []
    
    # Title
    elements.append(Paragraph("Academic Performance Report", title_style))
    elements.append(Spacer(1, 0.2*inch))
    
    # User info
    user_info = f"<b>Faculty:</b> {user.name} ({user.email})<br/><b>Generated:</b> {current_date}"
    elements.append(Paragraph(user_info, normal_style))
    elements.append(Spacer(1, 0.3*inch))
    
    # Executive Summary
    elements.append(Paragraph("Executive Summary", heading_style))
    
    summary_data = [["Metric", "Count"]]
    if summary.get('total_publications'):
        summary_data.append(["Publications", str(summary['total_publications'])])
    if summary.get('total_projects'):
        summary_data.append(["Research Projects", str(summary['total_projects'])])
    if summary.get('total_books'):
        summary_data.append(["Books & Chapters", str(summary['total_books'])])
    if summary.get('total_patents'):
        summary_data.append(["Patents", str(summary['total_patents'])])
    if summary.get('total_awards'):
        summary_data.append(["Awards", str(summary['total_awards'])])
    if summary.get('total_conferences'):
        summary_data.append(["Conferences", str(summary['total_conferences'])])
    if summary.get('total_fdp'):
        summary_data.append(["FDP Programs", str(summary['total_fdp'])])
    if summary.get('total_courses'):
        summary_data.append(["Courses", str(summary['total_courses'])])
    
    if len(summary_data) > 1:
        summary_table = Table(summary_data, colWidths=[3*inch, 1.5*inch])
        summary_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2563eb')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 11),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8fafc')])
        ]))
        elements.append(summary_table)
        elements.append(Spacer(1, 0.2*inch))
    
    # Analysis & Insights
    if interpretation:
        elements.append(Paragraph("Analysis & Insights", heading_style))
        # Strip HTML tags from interpretation
        clean_text = interpretation.replace('<div style=\'color:#475569;line-height:1.8;\'>', '').replace('</div>', '')
        clean_text = clean_text.replace('<p><strong>', '<b>').replace('</strong>', '</b>')
        clean_text = clean_text.replace('</p>', '<br/>')
        elements.append(Paragraph(clean_text, normal_style))
        elements.append(Spacer(1, 0.2*inch))
    
    # Detailed Statistics
    if detailed_stats:
        elements.append(PageBreak())
        elements.append(Paragraph("Detailed Statistics", heading_style))
        
        for category, stats in detailed_stats.items():
            elements.append(Paragraph(f"<b>{category}</b>", ParagraphStyle('SubHeading', parent=styles['Normal'], fontSize=11, textColor=colors.HexColor('#1e40af'))))
            stats_data = [["Metric", "Count"]]
            for stat_name, stat_value in stats.items():
                stats_data.append([stat_name, str(stat_value)])
            
            stats_table = Table(stats_data, colWidths=[3*inch, 1.5*inch])
            stats_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#e0f2fe')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#0c4a6e')),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#cbd5e1')),
                ('FONTSIZE', (0, 1), (-1, -1), 9),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f0f9ff')])
            ]))
            elements.append(stats_table)
            elements.append(Spacer(1, 0.1*inch))
    
    # Footer
    elements.append(Spacer(1, 0.3*inch))
    footer_text = f"<i>This report was automatically generated by Faculty MIS on {current_date}</i>"
    elements.append(Paragraph(footer_text, ParagraphStyle('Footer', parent=styles['Normal'], fontSize=8, textColor=colors.grey, alignment=TA_CENTER)))
    
    # Build PDF
    doc.build(elements)
    
    return pdf_path

