import os
import re
import openai
import PyPDF2
from PIL import Image, ImageDraw, ImageFont
from gtts import gTTS
from moviepy.editor import ImageClip, AudioFileClip, concatenate_videoclips

openai.api_key = ''

def extract_text_from_pdf(pdf_path):
    pdfReader = PyPDF2.PdfReader(open(pdf_path, 'rb'))
    text = ""
    for page in pdfReader.pages:
        text += page.extract_text()
    return text

def generate_slides_content(text):
    # Set your OpenAI API key
    # openai.api_key = os.getenv("OPENAI_API_KEY")
    
    prompt = f"Summarize the following text into slide titles and bullet points for a presentation. Format it as 'Slide X: Title' followed by bullet points:\n\n{text}"
    
    response = openai.ChatCompletion.create(
        model="gpt-4",
        messages=[
            {"role": "system", "content": "You are an assistant that creates presentation slides."},
            {"role": "user", "content": prompt}
        ],
        max_tokens=1000,
        temperature=0.7,
    )
    
    slides_content = response['choices'][0]['message']['content']
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

def generate_presentation_script(slide_content, original_text):
    # Use GPT-4 to generate a script for the slide
    prompt = f"Create a detailed and engaging presentation script for the following slide content based on the original text:\n\nSlide Content:\n{slide_content}\n\nOriginal Text:\n{original_text}"
    
    response = openai.ChatCompletion.create(
        model="gpt-4",
        messages=[
            {"role": "system", "content": "You are an assistant that writes presentation scripts."},
            {"role": "user", "content": prompt}
        ],
        max_tokens=1000,
        temperature=0.7,
    )
    
    script = response['choices'][0]['message']['content']
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
    
    # Step 2: Generate slides content
    print("Generating slides content...")
    slides_content = generate_slides_content(text)
    
    # Parse the slides content
    slides = parse_slides_content(slides_content)
    
    slide_filenames = []
    audio_filenames = []
    
    for idx, slide in enumerate(slides):
        print(f"Processing Slide {idx+1}: {slide['title']}")
        title = slide['title']
        bullet_points = slide['bullet_points']
        
        # Step 3: Create slide image
        slide_filename = create_slide_image(title, bullet_points, idx+1)
        slide_filenames.append(slide_filename)
        
        # Step 4: Generate presentation script
        slide_content = f"Title: {title}\nBullet Points:\n" + "\n".join(bullet_points)
        script = generate_presentation_script(slide_content, text)
        
        # Step 5: Generate audio
        audio_filename = generate_audio(script, idx+1)
        audio_filenames.append(audio_filename)
    
    # Step 6: Create video
    print("Creating video presentation...")
    create_video(slide_filenames, audio_filenames)
    print("Video presentation created successfully!")

if __name__ == '__main__':
    pdf_path = 'input.pdf'  # Replace with your PDF file path
    main(pdf_path)
