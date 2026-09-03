from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# Color Palette
NAVY = RGBColor(26, 35, 50)
IVORY = RGBColor(245, 241, 232)
BLUE = RGBColor(74, 111, 165)
TEAL = RGBColor(45, 138, 138)
AMBER = RGBColor(212, 165, 116)
WHITE = RGBColor(255, 255, 255)
LIGHT_GRAY = RGBColor(240, 238, 235)
RED = RGBColor(220, 53, 69)

def add_title_slide(prs, title, subtitle):
    """Add a title slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = NAVY
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    p = title_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(66)
    p.font.bold = True
    p.font.color.rgb = IVORY
    p.alignment = PP_ALIGN.CENTER
    
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(9), Inches(1))
    subtitle_frame = subtitle_box.text_frame
    p = subtitle_frame.paragraphs[0]
    p.text = subtitle
    p.font.size = Pt(36)
    p.font.color.rgb = AMBER
    p.alignment = PP_ALIGN.CENTER

def add_content_slide(prs, title, bg_color=NAVY):
    """Add a content slide with title"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = bg_color
    
    title_shape = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(0.9))
    title_shape.fill.solid()
    title_shape.fill.fore_color.rgb = TEAL
    title_shape.line.color.rgb = TEAL
    
    title_frame = title_shape.text_frame
    p = title_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = IVORY
    p.space_before = Pt(12)
    p.space_after = Pt(12)
    title_frame.margin_left = Inches(0.5)
    
    return slide

def add_text_box(slide, left, top, width, height, text, font_size=24, bold=False, color=NAVY, align=PP_ALIGN.LEFT):
    """Add a text box to a slide"""
    textbox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    text_frame = textbox.text_frame
    text_frame.word_wrap = True
    p = text_frame.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.alignment = align
    return textbox

def add_box(slide, left, top, width, height, text, font_size=14, bg_color=BLUE, text_color=IVORY, border_color=AMBER, bold=False):
    """Add a colored box with text"""
    shape = slide.shapes.add_shape(1, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    shape.line.color.rgb = border_color
    shape.line.width = Pt(2)
    
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = text_color
    p.alignment = PP_ALIGN.CENTER
    tf.vertical_anchor = 1
    return shape

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # ============ SLIDE 1: COVER ============
    add_title_slide(prs, "TravelGuard AI", "Predict. Detect. Recover.")
    
    # ============ SLIDE 2: THE PROBLEM ============
    slide = add_content_slide(prs, "The Problem: One Delay Breaks Everything")
    
    add_text_box(slide, 0.5, 1.2, 9, 0.6, 
                 "Modern travel is multimodal: Flight → Transfer → Train → Hotel", 
                 font_size=18, color=IVORY, bold=True)
    
    # Journey visualization
    journey = "Mumbai\n↓\n✈️ Flight (10:00-12:15)\n↓\n🚕 Transfer (60 min needed)\n↓\n🚆 Train (14:00 departure)\n↓\nJaipur"
    add_box(slide, 1.5, 1.9, 3, 3.5, journey, font_size=14, bg_color=BLUE, text_color=IVORY)
    
    # Disruption cascade
    cascade = "Flight Delayed +2h\n↓\nArrival: 14:30\n↓\nTrain Departs: 14:00\n↓\n❌ Connection Broken\n↓\n🔥 Entire Itinerary Fails"
    add_box(slide, 5.5, 1.9, 3, 3.5, cascade, font_size=14, bg_color=RED, text_color=IVORY, border_color=AMBER)
    
    add_text_box(slide, 0.8, 5.7, 8.4, 1.2, 
                 "Travelers manually check status, recalculate feasibility, search alternatives, rebuild itineraries.\nTravelGuard automates this entire process.",
                 font_size=14, color=IVORY)
    
    # ============ SLIDE 3: OUR SOLUTION ============
    slide = add_content_slide(prs, "Our Solution: Predict → Detect → Recover")
    
    add_text_box(slide, 0.5, 1.1, 9, 0.4, 
                 "A three-phase intelligent recovery system", 
                 font_size=16, color=AMBER, bold=True)
    
    # Three phases
    phases_data = [
        ("PREDICT\n\nML models forecast\ndelay probability\nfor each leg", 1.2, BLUE),
        ("DETECT\n\nNeo4j graph analyzes\nconnection feasibility\nand identifies breaks", 4, TEAL),
        ("RECOVER\n\nAI finds context-aware\nalternatives ranked\nby risk & cost", 6.8, BLUE)
    ]
    
    for text, left, color in phases_data:
        add_box(slide, left, 2, 2.5, 3.8, text, font_size=13, bg_color=color, text_color=IVORY, bold=True)
    
    # Flow arrows
    add_text_box(slide, 3.6, 3.8, 0.4, 0.5, "→", font_size=32, color=AMBER, align=PP_ALIGN.CENTER)
    add_text_box(slide, 6.4, 3.8, 0.4, 0.5, "→", font_size=32, color=AMBER, align=PP_ALIGN.CENTER)
    
    # ============ SLIDE 4: TECHNICAL ARCHITECTURE ============
    slide = add_content_slide(prs, "Technical Architecture: 5-Layer Stack")
    
    layers = [
        ("FRONTEND", "React Dashboard", BLUE, 1.2),
        ("BACKEND", "Node.js + Express", TEAL, 2.2),
        ("DATA LAYER", "MongoDB + Neo4j Graph DB", BLUE, 3.2),
        ("ML SERVICE", "Python + FastAPI", TEAL, 4.2),
        ("AI ORCHESTRATION", "Groq LLM Agent", BLUE, 5.2)
    ]
    
    for layer_name, tech, color, top in layers:
        layer_shape = slide.shapes.add_shape(1, Inches(1), Inches(top), Inches(8), Inches(0.7))
        layer_shape.fill.solid()
        layer_shape.fill.fore_color.rgb = color
        layer_shape.line.color.rgb = AMBER
        layer_shape.line.width = Pt(1)
        
        tf = layer_shape.text_frame
        p = tf.paragraphs[0]
        p.text = f"{layer_name}  →  {tech}"
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = IVORY
        p.space_before = Pt(4)
        tf.margin_left = Inches(0.3)
        
        if top < 5.2:
            add_text_box(slide, 4.7, top + 0.75, 0.6, 0.3, "↓", font_size=20, color=AMBER, align=PP_ALIGN.CENTER)
    
    # ============ SLIDE 5: DELAY PREDICTION + NEO4J ============
    slide = add_content_slide(prs, "Delay Prediction & Travel Graph Intelligence")
    
    # Left: ML Prediction
    add_text_box(slide, 0.5, 1.1, 4.5, 0.4, "ML Delay Prediction", font_size=14, bold=True, color=AMBER)
    
    pred_box = "Historical Data\n↓\nFeatures: Airline, Route,\nTime, Weather, Season\n↓\nModel: XGBoost\n↓\nOutput:\n✈️ Flight AI123\nDelay Probability: 82%\nRisk: HIGH"
    add_box(slide, 0.5, 1.6, 4.5, 4.2, pred_box, font_size=12, bg_color=BLUE, text_color=IVORY)
    
    # Right: Neo4j Graph
    add_text_box(slide, 5.2, 1.1, 4.3, 0.4, "Neo4j Travel Graph", font_size=14, bold=True, color=AMBER)
    
    graph_box = "Entities & Relationships:\n\nTraveler → Trip → Flight\n                ↓\n          Airport → Transfer\n                ↓\n            Train → Jaipur\n\nBenefit: Natural traversal\nof journey dependencies\nfor cascading detection"
    add_box(slide, 5.2, 1.6, 4.3, 4.2, graph_box, font_size=12, bg_color=TEAL, text_color=IVORY)
    
    # ============ SLIDE 6: DISRUPTION DETECTION IN ACTION ============
    slide = add_content_slide(prs, "Real-Time Disruption Detection")
    
    add_text_box(slide, 0.5, 1.1, 9, 0.4, 
                 "Automated: Graph analyzes live status against itinerary constraints", 
                 font_size=14, color=AMBER, bold=True)
    
    # Status input
    status = "LIVE STATUS:\n✈️ Flight AI123\nDELAYED +2h 15m\nEst. Arrival: 14:30"
    add_box(slide, 0.8, 1.7, 2.8, 2, status, font_size=12, bg_color=BLUE, text_color=IVORY)
    
    add_text_box(slide, 3.8, 2.6, 0.5, 0.5, "���", font_size=28, color=AMBER, align=PP_ALIGN.CENTER)
    
    # Graph check
    check = "NEO4J CHECK:\nFlight arrival: 14:30\nTrain departure: 14:00\nRequired transfer: 60 min\nAvailable time: -30 min"
    add_box(slide, 4.5, 1.7, 2.8, 2, check, font_size=12, bg_color=TEAL, text_color=IVORY)
    
    add_text_box(slide, 7.5, 2.6, 0.5, 0.5, "→", font_size=28, color=AMBER, align=PP_ALIGN.CENTER)
    
    # Result
    result = "❌ DISRUPTED\nConnection\nImpossible"
    add_box(slide, 8.2, 1.7, 1.5, 2, result, font_size=13, bg_color=RED, text_color=IVORY, bold=True)
    
    add_text_box(slide, 0.8, 4.1, 8.4, 2.8, 
                 "Recovery Trigger Activated:\nThe system automatically searches for alternatives and ranks them by connection reliability, cost, travel time, and delay risk. The traveler receives an intelligently ranked recovery plan within seconds.",
                 font_size=13, color=IVORY)
    
    # ============ SLIDE 7: RECOVERY ENGINE ============
    slide = add_content_slide(prs, "Intelligent Recovery: Alternative Ranking")
    
    add_text_box(slide, 0.5, 1.1, 9, 0.4, 
                 "Multi-criteria scoring: Connection Reliability 30% | Cost 20% | Time 25% | Risk 15% | Preference 10%", 
                 font_size=12, color=AMBER, bold=True)
    
    # Option A (Recommended)
    opt_a = "✅ RECOMMENDED\n\nFlight 14:40 + Train 18:00\n\n₹6,400\nLow Delay Risk\n1 Transfer\nArrive: 10:45 PM"
    add_box(slide, 0.8, 1.7, 2.8, 4, opt_a, font_size=11, bg_color=TEAL, text_color=IVORY, bold=True)
    
    # Option B
    opt_b = "Alternative\n\nDirect Train 16:30\n(Skip flight alternative)\n\n₹4,200\nMedium Risk\n0 Transfers\nArrive: 11:30 PM"
    add_box(slide, 4, 1.7, 2.8, 4, opt_b, font_size=11, bg_color=BLUE, text_color=IVORY)
    
    # Option C
    opt_c = "Context-Aware\n\nOvernight Hotel\n+ Morning Train 6:30 AM\n\n₹3,800 + Hotel\nLowest Risk\nRested & Fresh\nArrive: 7:00 AM"
    add_box(slide, 7.2, 1.7, 2.8, 4, opt_c, font_size=11, bg_color=BLUE, text_color=IVORY)
    
    # ============ SLIDE 8: AGENTIC AI + DEMO ============
    slide = add_content_slide(prs, "AI Agent Orchestration & Live Demo")
    
    add_text_box(slide, 0.5, 1.1, 4.5, 0.4, 
                 "Agent Coordinates Systems", 
                 font_size=13, bold=True, color=AMBER)
    
    agent_flow = """AI Recovery Agent
    
Connects:
• Transport Status
• ML Predictions
• Neo4j Graph
• Alternative Search
• User Preferences

Output:
Ranked recovery plan
with explanation"""
    
    add_box(slide, 0.5, 1.6, 4.5, 4.2, agent_flow, font_size=11, bg_color=TEAL, text_color=IVORY)
    
    # Demo timeline
    add_text_box(slide, 5.2, 1.1, 4.3, 0.4, 
                 "Live Walkthrough", 
                 font_size=13, bold=True, color=AMBER)
    
    demo = """T=0:  Flight on time
T=30: ML Alert: 82% delay risk
T=90: Flight delayed +2h 15m
T=120: Graph detects disruption
T=150: Recovery plan generated
T=180: User books Plan A
✅ Crisis averted"""
    
    add_box(slide, 5.2, 1.6, 4.3, 4.2, demo, font_size=11, bg_color=BLUE, text_color=IVORY)
    
    # ============ SLIDE 9: KEY DIFFERENTIATOR & CLOSE ============
    slide = add_content_slide(prs, "Why TravelGuard Wins")
    
    traditional = "Competitors:\n\nSEARCH → BOOK → TRACK\n\nNotify users of problems.\nUsers manually recover."
    add_box(slide, 0.8, 1.5, 4, 3, traditional, font_size=12, bg_color=LIGHT_GRAY, text_color=NAVY, bold=True)
    
    add_text_box(slide, 4.5, 2.8, 1, 0.5, "vs", font_size=20, bold=True, color=AMBER, align=PP_ALIGN.CENTER)
    
    travelguard = "TravelGuard:\n\nPLAN → PREDICT → DETECT\n→ RECOVER\n\nUnderstand dependencies.\nAutomatically recover.\nExplain decisions."
    add_box(slide, 5.2, 1.5, 4, 3, travelguard, font_size=12, bg_color=TEAL, text_color=IVORY, bold=True)
    
    add_text_box(slide, 0.8, 4.8, 8.4, 2.2, 
                 "\"TravelGuard doesn't just tell travelers something went wrong.\nIt understands how the disruption affects their entire journey and delivers an intelligent recovery plan.\"\n\n→ Deployed: Vercel + Render/Railway  |  →  Tech Stack: React, Node.js, Neo4j, FastAPI, Groq",
                 font_size=12, color=IVORY, align=PP_ALIGN.CENTER, bold=True)
    
    # Save
    prs.save('TravelGuard_AI_Presentation.pptx')
    print("✅ Professional 9-slide presentation created!")
    print("📊 File: TravelGuard_AI_Presentation.pptx")
    print("\n🎯 Slide Breakdown:")
    print("  1. Cover - Title & Tagline")
    print("  2. Problem - Cascading disruptions illustrated")
    print("  3. Solution - Predict → Detect → Recover")
    print("  4. Architecture - 5-layer tech stack")
    print("  5. Intelligence - ML + Neo4j capabilities")
    print("  6. Detection - Real-time logic in action")
    print("  7. Recovery - Multi-criteria ranking")
    print("  8. Orchestration - Agent + Live demo")
    print("  9. Differentiator - Why we win + deployment")

if __name__ == "__main__":
    create_presentation()
