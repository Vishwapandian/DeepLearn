import os
import re
import openai
import PyPDF2
from PIL import Image, ImageDraw, ImageFont
from gtts import gTTS
from moviepy.editor import ImageClip, AudioFileClip, concatenate_videoclips
from nltk.tokenize import sent_tokenize

import nltk
nltk.download('punkt')

openai.api_key = ''

def extract_text_from_pdf(pdf_path):
    pdfReader = PyPDF2.PdfReader(open(pdf_path, 'rb'))
    text = ""
    for page in pdfReader.pages:
        text += page.extract_text()
    return text

def summarize_text(text, max_tokens=500):
    # Summarize the text to reduce token usage
    prompt = f"Please provide a concise summary of the following text:\n\n{text}"
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",  # Use a less expensive model
        messages=[
            {"role": "system", "content": "You are an assistant that summarizes text efficiently."},
            {"role": "user", "content": prompt}
        ],
        max_tokens=max_tokens,
        temperature=0.5,
    )
    summary = response['choices'][0]['message']['content'].strip()
    return summary

def generate_slides_content(summarized_text):
    # Generate slides content using the summarized text
    prompt = f"Create up to 5 slide titles with bullet points from the following summary. Keep it concise:\n\n{summarized_text}"
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",  # Use a less expensive model
        messages=[
            {"role": "system", "content": "You are an assistant that creates concise presentation slides."},
            {"role": "user", "content": prompt}
        ],
        max_tokens=350,
        temperature=0.5,
    )
    slides_content = response['choices'][0]['message']['content'].strip()
    return slides_content

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

def create_slide_image(title, bullet_points, slide_number):
    # Define image size and colors
    img_width, img_height = 1280, 720
    background_color = 'white'
    title_color = 'black'
    bullet_color = 'black'
    
    # Create an image
    img = Image.new('RGB', (img_width, img_height), color=background_color)
    draw = ImageDraw.Draw(img)
    
    # Load fonts (ensure the font path is correct)
    try:
        title_font = ImageFont.truetype("arial.ttf", 60)
        bullet_font = ImageFont.truetype("arial.ttf", 40)
    except IOError:
        # If the font is not found, use a default font
        title_font = ImageFont.load_default()
        bullet_font = ImageFont.load_default()
    
    # Calculate positions
    title_position = (50, 50)
    bullet_start_y = 150
    line_spacing = 50
    
    # Draw title
    draw.text(title_position, title, font=title_font, fill=title_color)
    
    # Draw bullet points
    y_text = bullet_start_y
    for bullet in bullet_points:
        draw.text((70, y_text), f"• {bullet}", font=bullet_font, fill=bullet_color)
        y_text += line_spacing
    
    # Save image
    slide_filename = f"slide_{slide_number}.png"
    img.save(slide_filename)
    return slide_filename

def generate_presentation_script(slide_content, summarized_text):
    # Use the slide content and summarized text to generate a concise script
    prompt = f"Write a brief and engaging script for a presentation slide based on the following:\n\nSlide Title: {slide_content['title']}\nBullet Points:\n" + "\n".join(f"- {bp}" for bp in slide_content['bullet_points']) + f"\n\nUse the following summary for context:\n{summarized_text}"
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",  # Use less expensive model
        messages=[
            {"role": "system", "content": "You are an assistant that writes concise presentation scripts."},
            {"role": "user", "content": prompt}
        ],
        max_tokens=250,
        temperature=0.5,
    )
    script = response['choices'][0]['message']['content'].strip()
    return script

def generate_audio(script, slide_number, language='en'):
    tts = gTTS(text=script, lang=language)
    audio_filename = f"audio_{slide_number}.mp3"
    tts.save(audio_filename)
    return audio_filename

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

def main(pdf_path):
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
    
    slide_filenames = []
    audio_filenames = []
    
    for idx, slide in enumerate(slides):
        print(f"Processing Slide {idx+1}: {slide['title']}")
        
        # Step 4: Create slide image
        slide_filename = create_slide_image(slide['title'], slide['bullet_points'], idx+1)
        slide_filenames.append(slide_filename)
        
        # Step 5: Generate presentation script
        script = generate_presentation_script(slide, summarized_text)
        
        # Step 6: Generate audio
        audio_filename = generate_audio(script, idx+1)
        audio_filenames.append(audio_filename)
    
    # Step 7: Create video
    print("Creating video presentation...")
    create_video(slide_filenames, audio_filenames)
    print("Video presentation created successfully!")

if __name__ == '__main__':
    pdf_path = 'input.pdf'  # Replace with your PDF file path
    main(pdf_path)
