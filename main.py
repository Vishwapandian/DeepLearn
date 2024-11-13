import os
import re
import PyPDF2
from moviepy.editor import ImageClip, AudioFileClip, concatenate_videoclips
from nltk.tokenize import sent_tokenize
from pathlib import Path
from openai import OpenAI
import subprocess
from pdf2image import convert_from_path
import nltk
nltk.download('punkt')

# Import pptx library for creating PPTX files
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# Initialize OpenAI client with API key
client = OpenAI(
    api_key = ''
)

def extract_text_from_pdf(pdf_path):
    pdfReader = PyPDF2.PdfReader(open(pdf_path, 'rb'))
    text = ""
    for page in pdfReader.pages:
        text += page.extract_text()
    return text

def summarize_text(text, max_tokens=500):
    # Summarize the text to reduce token usage
    prompt = f"Please provide a concise summary of the following text:\n\n{text}"
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are an assistant that summarizes text efficiently."},
            {"role": "user", "content": prompt}
        ],
        max_tokens=max_tokens,
        temperature=0.5,
    )
    summary = response.choices[0].message.content.strip()
    print(summary)
    return summary

def generate_slides_content(summarized_text):
    # Generate slides content using the summarized text
    prompt = f"Create up to 5 slide titles with bullet points from the following summary. Keep it concise:\n\n{summarized_text}"
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are an assistant that creates concise presentation slides."},
            {"role": "user", "content": prompt}
        ],
        max_tokens=350,
        temperature=0.5,
    )
    slides_content = response.choices[0].message.content.strip()
    print(slides_content)
    return slides_content

def generate_audio(script, slide_number, voice="alloy"):
    """Generate audio using OpenAI's TTS API"""
    speech_file_path = Path(f"audio_{slide_number}.mp3")
    
    response = client.audio.speech.create(
        model="tts-1",
        voice=voice,
        input=script
    )
    
    # Save the audio file
    response.stream_to_file(str(speech_file_path))
    
    return str(speech_file_path)

def generate_presentation_script(slide_content, summarized_text):
    # Use the slide content and summarized text to generate a concise script
    prompt = f"Write a brief and engaging script for a presentation slide based on the following:\n\nSlide Title: {slide_content['title']}\nBullet Points:\n" + "\n".join(f"- {bp}" for bp in slide_content['bullet_points']) + f"\n\nUse the following summary for context:\n{summarized_text}"
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are an assistant that writes concise presentation scripts."},
            {"role": "user", "content": prompt}
        ],
        max_tokens=250,
        temperature=0.5,
    )
    script = response.choices[0].message.content.strip()
    print(script)
    return script

def parse_slides_content(slides_content):
    slides = []
    # Split slides content into individual slides
    slide_sections = re.split(r'\n(?=### Slide \d+:)', slides_content)
    for slide_section in slide_sections:
        # Extract title
        title_match = re.match(r'### Slide \d+: (.+)', slide_section)
        if title_match:
            title = title_match.group(1).strip()
            # Extract bullet points
            bullet_points = re.findall(r'- (.+)', slide_section)
            slides.append({'title': title, 'bullet_points': bullet_points})
    return slides

# New function to create intro and outro slides
def create_intro_slide(prs, title_text, subtitle_text):
    slide = prs.slides.add_slide(prs.slide_layouts[0])  # Title Slide layout
    title_placeholder = slide.shapes.title
    subtitle_placeholder = slide.placeholders[1]

    # Set title
    title_placeholder.text = title_text
    title_placeholder.text_frame.paragraphs[0].font.size = Pt(48)
    title_placeholder.text_frame.paragraphs[0].font.bold = True
    title_placeholder.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

    # Set subtitle
    subtitle_placeholder.text = subtitle_text
    subtitle_placeholder.text_frame.paragraphs[0].font.size = Pt(24)
    subtitle_placeholder.text_frame.paragraphs[0].font.color.rgb = RGBColor(200, 200, 200)

    # Set background color
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0, 70, 127)  # Dark blue

def create_outro_slide(prs, thank_you_text):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank Slide layout
    left = top = Inches(1)
    width = prs.slide_width - Inches(2)
    height = prs.slide_height - Inches(2)

    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    p = tf.add_paragraph()
    p.text = thank_you_text
    p.font.size = Pt(48)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = 1  # Center alignment

    # Set background color
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0, 70, 127)  # Dark blue

# Updated function to create a PPTX presentation with better visuals and intro/outro slides
def create_presentation(slides, pptx_filename='presentation.pptx'):
    prs = Presentation()
    # Set slide size to widescreen 16:9
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    # Create intro slide
    create_intro_slide(prs, "Presentation Title", "Subtitle or Presenter Name")

    for slide_content in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])  # Using Title and Content layout
        title_placeholder = slide.shapes.title
        content_placeholder = slide.placeholders[1]
        
        # Set the title
        title_placeholder.text = slide_content['title']
        title_placeholder.text_frame.paragraphs[0].font.size = Pt(36)
        title_placeholder.text_frame.paragraphs[0].font.bold = True
        title_placeholder.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 70, 127)

        # Add bullet points
        tf = content_placeholder.text_frame
        tf.clear()  # Clear any existing content
        for bullet_point in slide_content['bullet_points']:
            p = tf.add_paragraph()
            p.text = bullet_point
            p.level = 0
            p.font.size = Pt(24)
            p.font.color.rgb = RGBColor(50, 50, 50)
        
        # Set background color
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(230, 230, 230)  # Light gray background

    # Create outro slide
    create_outro_slide(prs, "Thank You!")

    # Save the presentation
    prs.save(pptx_filename)

# New function to export slides to images
def export_slides_to_images(pptx_filename, output_folder='slides'):
    # Create output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Define the paths
    pptx_path = os.path.abspath(pptx_filename)
    pdf_path = os.path.join(output_folder, "presentation.pdf")

    # Step 1: Convert PPTX to PDF
    command = [
        "soffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_folder,
        pptx_path
    ]
    subprocess.run(command, check=True)

    # Step 2: Convert PDF to individual PNG images
    slide_filenames = []
    images = convert_from_path(pdf_path)
    for i, image in enumerate(images):
        slide_filename = os.path.join(output_folder, f"slide_{i+1}.png")
        image.save(slide_filename, "PNG")
        slide_filenames.append(slide_filename)

    # Remove the intermediate PDF file if not needed
    os.remove(pdf_path)

    return slide_filenames

def create_video(slide_filenames, audio_filenames, output_filename="presentation.mp4"):
    clips = []
    for slide_filename, audio_filename in zip(slide_filenames, audio_filenames):
        # Get audio duration
        audio_clip = AudioFileClip(audio_filename)
        duration = audio_clip.duration

        # Create ImageClip with duration equal to audio duration
        image_clip = ImageClip(slide_filename).set_duration(duration)

        # Set audio
        image_clip = image_clip.set_audio(audio_clip)

        clips.append(image_clip)

    # Concatenate clips
    final_clip = concatenate_videoclips(clips, method="compose")

    # Write the video file
    final_clip.write_videofile(output_filename, fps=24)

    # Close clips to release resources
    final_clip.close()
    for clip in clips:
        clip.close()

def main(pdf_path, voice="alloy"):
    # Step 1: Extract text from PDF
    print("Extracting text from PDF...")
    text = extract_text_from_pdf(pdf_path)

    # Step 2: Summarize the text
    print("Summarizing text to reduce token usage...")
    summarized_text = summarize_text(text)

    # Step 3: Generate slides content
    print("Generating slides content...")
    slides_content = generate_slides_content(summarized_text)

    # Parse the slides content
    slides = parse_slides_content(slides_content)

    # List to track temp files
    temp_files = []

    # Step 4: Create PPTX presentation
    print("Creating PPTX presentation...")
    pptx_filename = 'presentation.pptx'
    create_presentation(slides, pptx_filename)
    temp_files.append(pptx_filename)  # Track the PPTX file

    # Step 5: Export slides to images
    print("Exporting slides to images...")
    slide_filenames = export_slides_to_images(pptx_filename)
    temp_files.extend(slide_filenames)  # Track all slide image files

    # Ensure the number of slide images matches the number of slides plus intro and outro
    expected_slides = len(slides) + 2  # Intro and Outro slides
    if len(slide_filenames) != expected_slides:
        print("Error: Number of slide images does not match expected number of slides.")
        return

    audio_filenames = []

    # Generate audio for intro slide
    intro_script = "Welcome to this presentation. Let's dive into the topic."
    audio_filename = generate_audio(intro_script, 1, voice=voice)
    audio_filenames.append(audio_filename)
    temp_files.append(audio_filename)

    # Generate audio for main slides
    for idx, slide in enumerate(slides):
        print(f"Processing Slide {idx+1}: {slide['title']}")

        # Step 6: Generate presentation script
        script = generate_presentation_script(slide, summarized_text)

        # Step 7: Generate audio
        audio_filename = generate_audio(script, idx+2, voice=voice)  # idx+2 because intro slide is first
        audio_filenames.append(audio_filename)
        temp_files.append(audio_filename)

    # Generate audio for outro slide
    outro_script = "Thank you for watching this presentation."
    audio_filename = generate_audio(outro_script, len(slides)+2, voice=voice)
    audio_filenames.append(audio_filename)
    temp_files.append(audio_filename)

    # Step 8: Create video
    print("Creating video presentation...")
    output_filename = "presentation.mp4"
    create_video(slide_filenames, audio_filenames, output_filename)
    print("Video presentation created successfully!")

    # Clean up temporary files, keeping only the final MP4
    print("Cleaning up temporary files...")
    for temp_file in temp_files:
        try:
            os.remove(temp_file)
        except OSError as e:
            print(f"Error deleting {temp_file}: {e}")

    print("Temporary files deleted. Only the final MP4 file remains.")

if __name__ == '__main__':
    pdf_path = 'input.pdf'  # Replace with your PDF file path
    voice = 'alloy'  # Can be: alloy, echo, fable, onyx, nova, or shimmer
    main(pdf_path, voice)
