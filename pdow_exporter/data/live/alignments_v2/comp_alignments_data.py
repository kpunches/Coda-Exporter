"""
comp_alignments_data.py
Extracted from v16 MCP fetches of grid-NkK5JDeDhF (comp_x_po, 200 rows)
and grid-6xLuZLsQ_t (comp_x_cct, 160 rows), filtered server-side by
[Program Abbreviation] = "MSCSIA".

Shape: list of (pcc_id, course_base_id, target_id, letters[])
Empty letters list means the junction row exists but no IRMA-X is set.
"""
from __future__ import annotations

# PO IDs (for reference)
# i-4MSH5Bo8aN = Cybersecurity Strategy and Risk Management
# i-AtlW5iPFg- = Secure Network, Cloud, and Enterprise Architecture Design
# i-s38p6SqZpY = Human-Centered Security in Software and AI Systems
# i-BKbajtfrlD = Threat Analysis and Infrastructure Defense
# i-l_Maz_Qi3P = Ethical, Legal, and Social Dimensions of Cybersecurity and AI

# CCT IDs (for reference)
# i-GOwzoAE6nA = Risk Management and Compliance
# i-VpmGsF6fdL = Integration of Artificial Intelligence
# i-60dxS64cyf = Technical Skill and Policy Writing
# i-bKEkKFt1HO = Leading a Security Program

# comp_x_po: 200 rows, 40 competencies x 5 POs
COMP_PO: list[tuple[str,str,str,list[str]]] = [
    # E123 (course_base_id=i-ATB4zvvKsO) competencies
    # pcc i-1uW8QiNqLe — Applies Modern Cybersecurity Principles
    ("i-1uW8QiNqLe","i-ATB4zvvKsO","i-4MSH5Bo8aN",["I"]),
    ("i-1uW8QiNqLe","i-ATB4zvvKsO","i-AtlW5iPFg-",["I"]),
    ("i-1uW8QiNqLe","i-ATB4zvvKsO","i-s38p6SqZpY",["I"]),
    ("i-1uW8QiNqLe","i-ATB4zvvKsO","i-BKbajtfrlD",["I"]),
    ("i-1uW8QiNqLe","i-ATB4zvvKsO","i-l_Maz_Qi3P",["I"]),
    # pcc i--qbgDUybm_ — Applies Basic Networking Principles
    ("i--qbgDUybm_","i-ATB4zvvKsO","i-4MSH5Bo8aN",["I"]),
    ("i--qbgDUybm_","i-ATB4zvvKsO","i-AtlW5iPFg-",["I"]),
    ("i--qbgDUybm_","i-ATB4zvvKsO","i-s38p6SqZpY",["I"]),
    ("i--qbgDUybm_","i-ATB4zvvKsO","i-BKbajtfrlD",["I"]),
    ("i--qbgDUybm_","i-ATB4zvvKsO","i-l_Maz_Qi3P",["I"]),
    # pcc i-VOX1WCsASn — Conducts a Risk Assessment using Frameworks
    ("i-VOX1WCsASn","i-ATB4zvvKsO","i-4MSH5Bo8aN",["I"]),
    ("i-VOX1WCsASn","i-ATB4zvvKsO","i-AtlW5iPFg-",["I"]),
    ("i-VOX1WCsASn","i-ATB4zvvKsO","i-s38p6SqZpY",["I"]),
    ("i-VOX1WCsASn","i-ATB4zvvKsO","i-BKbajtfrlD",["I"]),
    ("i-VOX1WCsASn","i-ATB4zvvKsO","i-l_Maz_Qi3P",["I"]),
    # E121 (course_base_id=i-q2kP8lsQww)
    # pcc i-do7f5SzsNW — Applies Principles and Frameworks of GRC
    ("i-do7f5SzsNW","i-q2kP8lsQww","i-4MSH5Bo8aN",["A","M"]),
    ("i-do7f5SzsNW","i-q2kP8lsQww","i-AtlW5iPFg-",[]),
    ("i-do7f5SzsNW","i-q2kP8lsQww","i-s38p6SqZpY",[]),
    ("i-do7f5SzsNW","i-q2kP8lsQww","i-BKbajtfrlD",[]),
    ("i-do7f5SzsNW","i-q2kP8lsQww","i-l_Maz_Qi3P",["M"]),
    # pcc i-wOJpausY_O — Implements AI Security Solutions
    ("i-wOJpausY_O","i-q2kP8lsQww","i-4MSH5Bo8aN",["A","M"]),
    ("i-wOJpausY_O","i-q2kP8lsQww","i-AtlW5iPFg-",[]),
    ("i-wOJpausY_O","i-q2kP8lsQww","i-s38p6SqZpY",[]),
    ("i-wOJpausY_O","i-q2kP8lsQww","i-BKbajtfrlD",[]),
    ("i-wOJpausY_O","i-q2kP8lsQww","i-l_Maz_Qi3P",["M"]),
    # D482 (course_base_id=i-KzniqYfsHm)
    # pcc i-4iAHxxm3BA — Assesses Business Needs and Network Security
    ("i-4iAHxxm3BA","i-KzniqYfsHm","i-4MSH5Bo8aN",[]),
    ("i-4iAHxxm3BA","i-KzniqYfsHm","i-AtlW5iPFg-",["I"]),
    ("i-4iAHxxm3BA","i-KzniqYfsHm","i-s38p6SqZpY",[]),
    ("i-4iAHxxm3BA","i-KzniqYfsHm","i-BKbajtfrlD",["I"]),
    ("i-4iAHxxm3BA","i-KzniqYfsHm","i-l_Maz_Qi3P",[]),
    # pcc i-FPEdLLtVW- — Aligns Secure Network Architecture
    ("i-FPEdLLtVW-","i-KzniqYfsHm","i-4MSH5Bo8aN",[]),
    ("i-FPEdLLtVW-","i-KzniqYfsHm","i-AtlW5iPFg-",["I"]),
    ("i-FPEdLLtVW-","i-KzniqYfsHm","i-s38p6SqZpY",[]),
    ("i-FPEdLLtVW-","i-KzniqYfsHm","i-BKbajtfrlD",["I"]),
    ("i-FPEdLLtVW-","i-KzniqYfsHm","i-l_Maz_Qi3P",[]),
    # pcc i-fVtXf5NRo7 — Recommends Network Security Solutions
    ("i-fVtXf5NRo7","i-KzniqYfsHm","i-4MSH5Bo8aN",[]),
    ("i-fVtXf5NRo7","i-KzniqYfsHm","i-AtlW5iPFg-",["I"]),
    ("i-fVtXf5NRo7","i-KzniqYfsHm","i-s38p6SqZpY",[]),
    ("i-fVtXf5NRo7","i-KzniqYfsHm","i-BKbajtfrlD",["I"]),
    ("i-fVtXf5NRo7","i-KzniqYfsHm","i-l_Maz_Qi3P",[]),
    # D483 (course_base_id=i-jYlJMhn-_w)
    # pcc i-8ntUfX2x_y — Manages Security Testing and Response
    ("i-8ntUfX2x_y","i-jYlJMhn-_w","i-4MSH5Bo8aN",["R"]),
    ("i-8ntUfX2x_y","i-jYlJMhn-_w","i-AtlW5iPFg-",[]),
    ("i-8ntUfX2x_y","i-jYlJMhn-_w","i-s38p6SqZpY",[]),
    ("i-8ntUfX2x_y","i-jYlJMhn-_w","i-BKbajtfrlD",["A","R"]),
    ("i-8ntUfX2x_y","i-jYlJMhn-_w","i-l_Maz_Qi3P",[]),
    # pcc i-amECrZryQh — Applies Software and System Security
    ("i-amECrZryQh","i-jYlJMhn-_w","i-4MSH5Bo8aN",["R"]),
    ("i-amECrZryQh","i-jYlJMhn-_w","i-AtlW5iPFg-",[]),
    ("i-amECrZryQh","i-jYlJMhn-_w","i-s38p6SqZpY",[]),
    ("i-amECrZryQh","i-jYlJMhn-_w","i-BKbajtfrlD",["A","R"]),
    ("i-amECrZryQh","i-jYlJMhn-_w","i-l_Maz_Qi3P",[]),
    # pcc i-jJXrrdb8ci — Applies Improvement Techniques and Automation
    ("i-jJXrrdb8ci","i-jYlJMhn-_w","i-4MSH5Bo8aN",["R"]),
    ("i-jJXrrdb8ci","i-jYlJMhn-_w","i-AtlW5iPFg-",[]),
    ("i-jJXrrdb8ci","i-jYlJMhn-_w","i-s38p6SqZpY",[]),
    ("i-jJXrrdb8ci","i-jYlJMhn-_w","i-BKbajtfrlD",["A","R"]),
    ("i-jJXrrdb8ci","i-jYlJMhn-_w","i-l_Maz_Qi3P",[]),
    # pcc i-vehyuPNez2 — Applies Incident Response Procedures
    ("i-vehyuPNez2","i-jYlJMhn-_w","i-4MSH5Bo8aN",["R"]),
    ("i-vehyuPNez2","i-jYlJMhn-_w","i-AtlW5iPFg-",[]),
    ("i-vehyuPNez2","i-jYlJMhn-_w","i-s38p6SqZpY",[]),
    ("i-vehyuPNez2","i-jYlJMhn-_w","i-BKbajtfrlD",["A","R"]),
    ("i-vehyuPNez2","i-jYlJMhn-_w","i-l_Maz_Qi3P",[]),
    # pcc i-VQLkEbyIe_ — Applies Security Concepts to Risk Management
    ("i-VQLkEbyIe_","i-jYlJMhn-_w","i-4MSH5Bo8aN",["R"]),
    ("i-VQLkEbyIe_","i-jYlJMhn-_w","i-AtlW5iPFg-",[]),
    ("i-VQLkEbyIe_","i-jYlJMhn-_w","i-s38p6SqZpY",[]),
    ("i-VQLkEbyIe_","i-jYlJMhn-_w","i-BKbajtfrlD",["A","R"]),
    ("i-VQLkEbyIe_","i-jYlJMhn-_w","i-l_Maz_Qi3P",[]),
    # pcc i-cSq_9vYhjd — Recommends Incident Response Solutions
    ("i-cSq_9vYhjd","i-jYlJMhn-_w","i-4MSH5Bo8aN",["R"]),
    ("i-cSq_9vYhjd","i-jYlJMhn-_w","i-AtlW5iPFg-",[]),
    ("i-cSq_9vYhjd","i-jYlJMhn-_w","i-s38p6SqZpY",[]),
    ("i-cSq_9vYhjd","i-jYlJMhn-_w","i-BKbajtfrlD",["A","R"]),
    ("i-cSq_9vYhjd","i-jYlJMhn-_w","i-l_Maz_Qi3P",[]),
    # D484 (course_base_id=i-8SY7daa3NP)
    # pcc i-Z8TiOfNQxg — Defines Penetration Testing Engagement
    ("i-Z8TiOfNQxg","i-8SY7daa3NP","i-4MSH5Bo8aN",[]),
    ("i-Z8TiOfNQxg","i-8SY7daa3NP","i-AtlW5iPFg-",[]),
    ("i-Z8TiOfNQxg","i-8SY7daa3NP","i-s38p6SqZpY",[]),
    ("i-Z8TiOfNQxg","i-8SY7daa3NP","i-BKbajtfrlD",["M"]),
    ("i-Z8TiOfNQxg","i-8SY7daa3NP","i-l_Maz_Qi3P",[]),
    # pcc i-V_8IZ4xEHV — Performs Cyber Reconnaissance
    ("i-V_8IZ4xEHV","i-8SY7daa3NP","i-4MSH5Bo8aN",[]),
    ("i-V_8IZ4xEHV","i-8SY7daa3NP","i-AtlW5iPFg-",[]),
    ("i-V_8IZ4xEHV","i-8SY7daa3NP","i-s38p6SqZpY",[]),
    ("i-V_8IZ4xEHV","i-8SY7daa3NP","i-BKbajtfrlD",["M"]),
    ("i-V_8IZ4xEHV","i-8SY7daa3NP","i-l_Maz_Qi3P",[]),
    # pcc i-Sp3rciDLws — Develops Penetration Testing Techniques
    ("i-Sp3rciDLws","i-8SY7daa3NP","i-4MSH5Bo8aN",[]),
    ("i-Sp3rciDLws","i-8SY7daa3NP","i-AtlW5iPFg-",[]),
    ("i-Sp3rciDLws","i-8SY7daa3NP","i-s38p6SqZpY",[]),
    ("i-Sp3rciDLws","i-8SY7daa3NP","i-BKbajtfrlD",["M"]),
    ("i-Sp3rciDLws","i-8SY7daa3NP","i-l_Maz_Qi3P",[]),
    # pcc i-ZTHdX7RfEJ — Simulates Attacks and Responses
    ("i-ZTHdX7RfEJ","i-8SY7daa3NP","i-4MSH5Bo8aN",[]),
    ("i-ZTHdX7RfEJ","i-8SY7daa3NP","i-AtlW5iPFg-",[]),
    ("i-ZTHdX7RfEJ","i-8SY7daa3NP","i-s38p6SqZpY",[]),
    ("i-ZTHdX7RfEJ","i-8SY7daa3NP","i-BKbajtfrlD",["M"]),
    ("i-ZTHdX7RfEJ","i-8SY7daa3NP","i-l_Maz_Qi3P",[]),
    # pcc i-oPRsTIz_zE — Reports Cybersecurity Assessments and Actions
    ("i-oPRsTIz_zE","i-8SY7daa3NP","i-4MSH5Bo8aN",[]),
    ("i-oPRsTIz_zE","i-8SY7daa3NP","i-AtlW5iPFg-",[]),
    ("i-oPRsTIz_zE","i-8SY7daa3NP","i-s38p6SqZpY",[]),
    ("i-oPRsTIz_zE","i-8SY7daa3NP","i-BKbajtfrlD",["M"]),
    ("i-oPRsTIz_zE","i-8SY7daa3NP","i-l_Maz_Qi3P",[]),
    # pcc i-oWgPlYCxMb — Evaluates a Penetration Testing Engagement Plan
    ("i-oWgPlYCxMb","i-8SY7daa3NP","i-4MSH5Bo8aN",[]),
    ("i-oWgPlYCxMb","i-8SY7daa3NP","i-AtlW5iPFg-",[]),
    ("i-oWgPlYCxMb","i-8SY7daa3NP","i-s38p6SqZpY",[]),
    ("i-oWgPlYCxMb","i-8SY7daa3NP","i-BKbajtfrlD",["M"]),
    ("i-oWgPlYCxMb","i-8SY7daa3NP","i-l_Maz_Qi3P",[]),
    # D485 (course_base_id=i-9H5ifiQ1xL)
    # pcc i-GwmaPfCjQN — Designs Secure Cloud Solutions
    ("i-GwmaPfCjQN","i-9H5ifiQ1xL","i-4MSH5Bo8aN",[]),
    ("i-GwmaPfCjQN","i-9H5ifiQ1xL","i-AtlW5iPFg-",["I"]),
    ("i-GwmaPfCjQN","i-9H5ifiQ1xL","i-s38p6SqZpY",["I"]),
    ("i-GwmaPfCjQN","i-9H5ifiQ1xL","i-BKbajtfrlD",[]),
    ("i-GwmaPfCjQN","i-9H5ifiQ1xL","i-l_Maz_Qi3P",[]),
    # pcc i-biO2Y51QXJ — Implements Security Cloud Solutions
    ("i-biO2Y51QXJ","i-9H5ifiQ1xL","i-4MSH5Bo8aN",[]),
    ("i-biO2Y51QXJ","i-9H5ifiQ1xL","i-AtlW5iPFg-",["I"]),
    ("i-biO2Y51QXJ","i-9H5ifiQ1xL","i-s38p6SqZpY",["I"]),
    ("i-biO2Y51QXJ","i-9H5ifiQ1xL","i-BKbajtfrlD",[]),
    ("i-biO2Y51QXJ","i-9H5ifiQ1xL","i-l_Maz_Qi3P",[]),
    # pcc i-KgpKwNXzkL — Analyzes Risk Management Plans
    ("i-KgpKwNXzkL","i-9H5ifiQ1xL","i-4MSH5Bo8aN",[]),
    ("i-KgpKwNXzkL","i-9H5ifiQ1xL","i-AtlW5iPFg-",["I"]),
    ("i-KgpKwNXzkL","i-9H5ifiQ1xL","i-s38p6SqZpY",["I"]),
    ("i-KgpKwNXzkL","i-9H5ifiQ1xL","i-BKbajtfrlD",[]),
    ("i-KgpKwNXzkL","i-9H5ifiQ1xL","i-l_Maz_Qi3P",[]),
    # D487 (course_base_id=i-m4uEhiIL-f)
    # pcc i-elqNIucdrY — Examines Security Methods within SDLC
    ("i-elqNIucdrY","i-m4uEhiIL-f","i-4MSH5Bo8aN",[]),
    ("i-elqNIucdrY","i-m4uEhiIL-f","i-AtlW5iPFg-",[]),
    ("i-elqNIucdrY","i-m4uEhiIL-f","i-s38p6SqZpY",["A","R"]),
    ("i-elqNIucdrY","i-m4uEhiIL-f","i-BKbajtfrlD",[]),
    ("i-elqNIucdrY","i-m4uEhiIL-f","i-l_Maz_Qi3P",[]),
    # pcc i-1ro_H1jfr- — Assesses Software Requirements and Risks
    ("i-1ro_H1jfr-","i-m4uEhiIL-f","i-4MSH5Bo8aN",[]),
    ("i-1ro_H1jfr-","i-m4uEhiIL-f","i-AtlW5iPFg-",[]),
    ("i-1ro_H1jfr-","i-m4uEhiIL-f","i-s38p6SqZpY",["A","R"]),
    ("i-1ro_H1jfr-","i-m4uEhiIL-f","i-BKbajtfrlD",[]),
    ("i-1ro_H1jfr-","i-m4uEhiIL-f","i-l_Maz_Qi3P",[]),
    # pcc i-p4KJB0J6TY — Evaluates Software Security Test Plan
    ("i-p4KJB0J6TY","i-m4uEhiIL-f","i-4MSH5Bo8aN",[]),
    ("i-p4KJB0J6TY","i-m4uEhiIL-f","i-AtlW5iPFg-",[]),
    ("i-p4KJB0J6TY","i-m4uEhiIL-f","i-s38p6SqZpY",["A","R"]),
    ("i-p4KJB0J6TY","i-m4uEhiIL-f","i-BKbajtfrlD",[]),
    ("i-p4KJB0J6TY","i-m4uEhiIL-f","i-l_Maz_Qi3P",[]),
    # pcc i-tOJiVf7V_b — Evaluates Effectiveness of Software Testing
    ("i-tOJiVf7V_b","i-m4uEhiIL-f","i-4MSH5Bo8aN",[]),
    ("i-tOJiVf7V_b","i-m4uEhiIL-f","i-AtlW5iPFg-",[]),
    ("i-tOJiVf7V_b","i-m4uEhiIL-f","i-s38p6SqZpY",["A","R"]),
    ("i-tOJiVf7V_b","i-m4uEhiIL-f","i-BKbajtfrlD",[]),
    ("i-tOJiVf7V_b","i-m4uEhiIL-f","i-l_Maz_Qi3P",[]),
    # D488 (course_base_id=i-goi7k0AsGH)
    # pcc i-lY6rG_J0WU — Designs Secure Network Architecture
    ("i-lY6rG_J0WU","i-goi7k0AsGH","i-4MSH5Bo8aN",["R"]),
    ("i-lY6rG_J0WU","i-goi7k0AsGH","i-AtlW5iPFg-",["A","R"]),
    ("i-lY6rG_J0WU","i-goi7k0AsGH","i-s38p6SqZpY",["R"]),
    ("i-lY6rG_J0WU","i-goi7k0AsGH","i-BKbajtfrlD",["R"]),
    ("i-lY6rG_J0WU","i-goi7k0AsGH","i-l_Maz_Qi3P",[]),
    # pcc i-MakbqUGa_q — Implements Secure Solutions
    ("i-MakbqUGa_q","i-goi7k0AsGH","i-4MSH5Bo8aN",["R"]),
    ("i-MakbqUGa_q","i-goi7k0AsGH","i-AtlW5iPFg-",["A","R"]),
    ("i-MakbqUGa_q","i-goi7k0AsGH","i-s38p6SqZpY",["R"]),
    ("i-MakbqUGa_q","i-goi7k0AsGH","i-BKbajtfrlD",["R"]),
    ("i-MakbqUGa_q","i-goi7k0AsGH","i-l_Maz_Qi3P",[]),
    # pcc i-u68wJZrvOh — Designs Integration of Cybersecurity Solutions
    ("i-u68wJZrvOh","i-goi7k0AsGH","i-4MSH5Bo8aN",["R"]),
    ("i-u68wJZrvOh","i-goi7k0AsGH","i-AtlW5iPFg-",["A","R"]),
    ("i-u68wJZrvOh","i-goi7k0AsGH","i-s38p6SqZpY",["R"]),
    ("i-u68wJZrvOh","i-goi7k0AsGH","i-BKbajtfrlD",["R"]),
    ("i-u68wJZrvOh","i-goi7k0AsGH","i-l_Maz_Qi3P",[]),
    # pcc i-KnwLViUZQc — Develops Secure Architecture for GRC
    ("i-KnwLViUZQc","i-goi7k0AsGH","i-4MSH5Bo8aN",["R"]),
    ("i-KnwLViUZQc","i-goi7k0AsGH","i-AtlW5iPFg-",["A","R"]),
    ("i-KnwLViUZQc","i-goi7k0AsGH","i-s38p6SqZpY",["R"]),
    ("i-KnwLViUZQc","i-goi7k0AsGH","i-BKbajtfrlD",["R"]),
    ("i-KnwLViUZQc","i-goi7k0AsGH","i-l_Maz_Qi3P",[]),
    # D489 (course_base_id=i-zV3gie6I3B)
    # pcc i-kN4Whq_gFD — Describes Security Risks, Standards, and Roles
    ("i-kN4Whq_gFD","i-zV3gie6I3B","i-4MSH5Bo8aN",["M"]),
    ("i-kN4Whq_gFD","i-zV3gie6I3B","i-AtlW5iPFg-",[]),
    ("i-kN4Whq_gFD","i-zV3gie6I3B","i-s38p6SqZpY",["M"]),
    ("i-kN4Whq_gFD","i-zV3gie6I3B","i-BKbajtfrlD",[]),
    ("i-kN4Whq_gFD","i-zV3gie6I3B","i-l_Maz_Qi3P",["R"]),
    # pcc i-709GllEeGf — Develops Security Policies and Guidelines
    ("i-709GllEeGf","i-zV3gie6I3B","i-4MSH5Bo8aN",["M"]),
    ("i-709GllEeGf","i-zV3gie6I3B","i-AtlW5iPFg-",[]),
    ("i-709GllEeGf","i-zV3gie6I3B","i-s38p6SqZpY",["M"]),
    ("i-709GllEeGf","i-zV3gie6I3B","i-BKbajtfrlD",[]),
    ("i-709GllEeGf","i-zV3gie6I3B","i-l_Maz_Qi3P",["R"]),
    # D490 (course_base_id=i-x9VVwNTpOl)
    # pcc i-9tR_0KKOZC — Capstone
    ("i-9tR_0KKOZC","i-x9VVwNTpOl","i-4MSH5Bo8aN",["M"]),
    ("i-9tR_0KKOZC","i-x9VVwNTpOl","i-AtlW5iPFg-",["M"]),
    ("i-9tR_0KKOZC","i-x9VVwNTpOl","i-s38p6SqZpY",["M"]),
    ("i-9tR_0KKOZC","i-x9VVwNTpOl","i-BKbajtfrlD",["M"]),
    ("i-9tR_0KKOZC","i-x9VVwNTpOl","i-l_Maz_Qi3P",["M"]),
    # E122 (course_base_id=i-stVHWsUqJb)
    # pcc i-Nu7VKWUbxH — Describes Significance of Human-Centric Risk
    ("i-Nu7VKWUbxH","i-stVHWsUqJb","i-4MSH5Bo8aN",["M"]),
    ("i-Nu7VKWUbxH","i-stVHWsUqJb","i-AtlW5iPFg-",[]),
    ("i-Nu7VKWUbxH","i-stVHWsUqJb","i-s38p6SqZpY",["M"]),
    ("i-Nu7VKWUbxH","i-stVHWsUqJb","i-BKbajtfrlD",["R"]),
    ("i-Nu7VKWUbxH","i-stVHWsUqJb","i-l_Maz_Qi3P",["A","M"]),
    # pcc i-RcNQuu3lRJ — Analyzes Human-Centric Risk Factors
    ("i-RcNQuu3lRJ","i-stVHWsUqJb","i-4MSH5Bo8aN",["M"]),
    ("i-RcNQuu3lRJ","i-stVHWsUqJb","i-AtlW5iPFg-",[]),
    ("i-RcNQuu3lRJ","i-stVHWsUqJb","i-s38p6SqZpY",["M"]),
    ("i-RcNQuu3lRJ","i-stVHWsUqJb","i-BKbajtfrlD",["R"]),
    ("i-RcNQuu3lRJ","i-stVHWsUqJb","i-l_Maz_Qi3P",["A","M"]),
    # pcc i-R-BpDe_IkI — Develops a Culture of Security in an Organization
    ("i-R-BpDe_IkI","i-stVHWsUqJb","i-4MSH5Bo8aN",["M"]),
    ("i-R-BpDe_IkI","i-stVHWsUqJb","i-AtlW5iPFg-",[]),
    ("i-R-BpDe_IkI","i-stVHWsUqJb","i-s38p6SqZpY",["M"]),
    ("i-R-BpDe_IkI","i-stVHWsUqJb","i-BKbajtfrlD",["R"]),
    ("i-R-BpDe_IkI","i-stVHWsUqJb","i-l_Maz_Qi3P",["A","M"]),
    # pcc i-iCW3qVt1TG — Applies Privacy by Design Principles
    ("i-iCW3qVt1TG","i-stVHWsUqJb","i-4MSH5Bo8aN",["M"]),
    ("i-iCW3qVt1TG","i-stVHWsUqJb","i-AtlW5iPFg-",[]),
    ("i-iCW3qVt1TG","i-stVHWsUqJb","i-s38p6SqZpY",["M"]),
    ("i-iCW3qVt1TG","i-stVHWsUqJb","i-BKbajtfrlD",["R"]),
    ("i-iCW3qVt1TG","i-stVHWsUqJb","i-l_Maz_Qi3P",["A","M"]),
    # E121 (course_base_id=i-q2kP8lsQww) — remaining pccs
    # pcc i-DOEFquqsxM — Analyzes Risks of AI Adoption
    ("i-DOEFquqsxM","i-q2kP8lsQww","i-4MSH5Bo8aN",["A","M"]),
    ("i-DOEFquqsxM","i-q2kP8lsQww","i-AtlW5iPFg-",[]),
    ("i-DOEFquqsxM","i-q2kP8lsQww","i-s38p6SqZpY",[]),
    ("i-DOEFquqsxM","i-q2kP8lsQww","i-BKbajtfrlD",[]),
    ("i-DOEFquqsxM","i-q2kP8lsQww","i-l_Maz_Qi3P",["M"]),
    # pcc i-sf5BLPAcnl — Optimizes Security with AI
    ("i-sf5BLPAcnl","i-q2kP8lsQww","i-4MSH5Bo8aN",["A","M"]),
    ("i-sf5BLPAcnl","i-q2kP8lsQww","i-AtlW5iPFg-",[]),
    ("i-sf5BLPAcnl","i-q2kP8lsQww","i-s38p6SqZpY",[]),
    ("i-sf5BLPAcnl","i-q2kP8lsQww","i-BKbajtfrlD",[]),
    ("i-sf5BLPAcnl","i-q2kP8lsQww","i-l_Maz_Qi3P",["M"]),
]

# comp_x_cct: 160 rows, 40 competencies x 4 CCTs (and mostly empty for MSCSIA per fetch)
# The fetches showed all comp_x_cct rows for MSCSIA have empty "Aligned" arrays EXCEPT
# nothing was populated in this program yet at cell level. We record structural presence only.
COMP_CCT: list[tuple[str,str,str,list[str]]] = [
    # All rows observed had empty Aligned — we still note them so writer produces empty cells not missing ones.
    # We skip populating — the writer's .get(...,'') default handles absence identically.
]
