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
        model="gpt-3.5-turbo-0125",
        messages=[
            {"role": "system", "content": "You are an assistant that summarizes text efficiently."},
            {"role": "user", "content": prompt}
        ],
        max_tokens=max_tokens,
        temperature=0.5,
    )
    summary = response.choices[0].message.content.strip()
    return summary

def generate_slides_content(summarized_text):
    # Generate slides content using the summarized text
    prompt = f"Create up to 5 slide titles with bullet points from the following summary. Keep it concise:\n\n{summarized_text}"
    response = client.chat.completions.create(
        model="gpt-3.5-turbo-0125",
        messages=[
            {"role": "system", "content": "You are an assistant that creates concise presentation slides."},
            {"role": "user", "content": prompt}
        ],
        max_tokens=350,
        temperature=0.5,
    )
    slides_content = response.choices[0].message.content.strip()
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
        model="gpt-3.5-turbo-0125",
        messages=[
            {"role": "system", "content": "You are an assistant that writes concise presentation scripts."},
            {"role": "user", "content": prompt}
        ],
        max_tokens=250,
        temperature=0.5,
    )
    script = response.choices[0].message.content.strip()
    return script

def parse_slides_content(slides_content):
    slides = []
    # Split slides content into individual slides
    slide_sections = re.split(r'\n(?=Slide \d+:)', slides_content)
    for slide_section in slide_sections:
        # Extract title
        title_match = re.match(r'Slide \d+: (.+)', slide_section)
        if title_match:
            title = title_match.group(1).strip()
            # Extract bullet points
            bullet_points = re.findall(r'- (.+)', slide_section)
            slides.append({'title': title, 'bullet_points': bullet_points})
    return slides

# New function to create a PPTX presentation
def create_presentation(slides, pptx_filename='presentation.pptx'):
    prs = Presentation()
    # Set slide width and height if needed
    # prs.slide_width = Inches(16)
    # prs.slide_height = Inches(9)

    for slide_content in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])  # Using Title and Content layout
        title_placeholder = slide.shapes.title
        content_placeholder = slide.placeholders[1]
        
        # Set the title
        title_placeholder.text = slide_content['title']
        
        # Add bullet points
        tf = content_placeholder.text_frame
        tf.clear()  # Clear any existing content
        for bullet_point in slide_content['bullet_points']:
            p = tf.add_paragraph()
            p.text = bullet_point
            p.level = 0  # Set bullet level if needed
            p.font.size = Pt(24)  # Set font size
        
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

    audio_filenames = []

    # Step 4: Create PPTX presentation
    print("Creating PPTX presentation...")
    pptx_filename = 'presentation.pptx'
    create_presentation(slides, pptx_filename)

    # Step 5: Export slides to images
    print("Exporting slides to images...")
    slide_filenames = export_slides_to_images(pptx_filename)

    # Ensure the number of slide images matches the number of slides
    if len(slide_filenames) != len(slides):
        print("Error: Number of slide images does not match number of slides.")
        return

    for idx, slide in enumerate(slides):
        print(f"Processing Slide {idx+1}: {slide['title']}")

        # Step 6: Generate presentation script
        script = generate_presentation_script(slide, summarized_text)

        # Step 7: Generate audio
        audio_filename = generate_audio(script, idx+1, voice=voice)
        audio_filenames.append(audio_filename)

    # Step 8: Create video
    print("Creating video presentation...")
    create_video(slide_filenames, audio_filenames)
    print("Video presentation created successfully!")

if __name__ == '__main__':
    pdf_path = 'input.pdf'  # Replace with your PDF file path
    voice = 'alloy'  # Can be: alloy, echo, fable, onyx, nova, or shimmer
    main(pdf_path, voice)
