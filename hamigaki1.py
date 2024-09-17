import cv2
import mediapipe as mp
import openai
import asyncio
import time
import threading
import tkinter as tk
from tkinter import Label, Canvas
from PIL import Image, ImageTk, ImageDraw, ImageFont, ImageEnhance, ImageChops
import pygame

pygame.mixer.init()
pygame.mixer.music.load('maou_bgm_piano25.mp3')

openai.api_key = 'sk-ZRZIHt2q59TkOqlz1alHCUMoW2B-8Wrk9q0iUeAlsVT3BlbkFJBqKrCi3Xt2HfG8Utc-Bwwk5NEmtHTwFHS5Dmpeg54A'

mp_hands = mp.solutions.hands
mp_face_mesh = mp.solutions.face_mesh

score = 0
start_time = 0
game_duration = 180
is_running = False
previous_hand_position = None
last_frame_time = 0
feedback_text = ""

cap = cv2.VideoCapture(0)

async def chatgpt_feedback(prompt):
    response = await openai.ChatCompletion.acreate(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content": "You are an assistant for analyzing toothbrushing motion."},
            {"role": "user", "content": prompt}
        ]
    )
    return response['choices'][0]['message']['content'].strip()

def call_chatgpt_in_background(prompt, rank_message):
    def run_in_thread():
        advice = asyncio.run(chatgpt_feedback(prompt))
        root.after(0, show_advice, rank_message, advice)
    thread = threading.Thread(target=run_in_thread)
    thread.start()

def show_advice(rank_message, advice):
    advice_label.config(text=f"{rank_message}\n{advice}")

def analyze_movement(hand_landmarks, mouth_landmarks, frame_width):
    global score, previous_hand_position, feedback_text
    feedback = "画面上に手または顔がありません。"
    current_time = time.time()

    if hand_landmarks and mouth_landmarks:
        current_hand_position = hand_landmarks[0]
        hand_x = current_hand_position[0] * frame_width

        if previous_hand_position is not None:
            movement_x = abs(current_hand_position[0] - previous_hand_position[0])
            movement_y = abs(current_hand_position[1] - previous_hand_position[1])
            movement_z = abs(current_hand_position[2] - previous_hand_position[2])

            if movement_x > 0.01 or movement_y > 0.01 or movement_z > 0.01:
                feedback = "正しい動きです。"
                score += 10
            else:
                feedback = "もっと磨きましょう。"
        previous_hand_position = current_hand_position
    feedback_text = feedback

    return feedback

def rank_judgment():
    global score
    if score > 26000:
        rank = "SS"
        advice_prompt = "ユーザーは最上級のSSランクを獲得しました。ブラッシングの動きを改善するためのアドバイスを超簡潔に提供してください。"
        rank_message = "3分経過！あなたのランクはSS！"
    elif score > 20000:
        rank = "S"
        advice_prompt = "ユーザーはSランクを獲得しました。ブラッシングの動きを改善するためのアドバイスを超簡潔に提供してください。"
        rank_message = "3分経過！あなたのランクはS！"
    elif score > 15000:
        rank = "A"
        advice_prompt = "ユーザーはAランクを獲得しました。ブラッシングの動きを改善するためのアドバイスを超簡潔に提供してください。"
        rank_message = "3分経過！あなたのランクはA！"
    elif score > 10000:
        rank = "B"
        advice_prompt = "ユーザーはBランクを獲得しました。ブラッシングの動きを改善するためのアドバイスを超簡潔に提供してください。"
        rank_message = "3分経過！あなたのランクはB！"
    else:
        rank = "C"
        advice_prompt = "ユーザーはCランクを獲得しました。ブラッシングの動きを改善するためのアドバイスを超簡潔に提供してください。"
        rank_message = "3分経過！あなたのランクはC！"
    
    call_chatgpt_in_background(advice_prompt, rank_message)

def start_game():
    global is_running, start_time, score
    is_running = True
    pygame.mixer.music.load('maou_bgm_piano25.mp3')
    score = 0
    start_time = time.time()
    pygame.mixer.music.play(-1)
    run_game()

def check_end_game():
    global is_running
    if is_running and time.time() - start_time >= game_duration:
        rank_judgment()
        is_running = False
        stop_game()
        pygame.mixer.music.load('ホイッスル音.mp3')
        pygame.mixer.music.play()

def run_game():
    global last_frame_time, feedback_text
    if is_running:
        ret, frame = cap.read()
        if not ret:
            return

        current_time = time.time()
        elapsed_time = current_time - start_time
        remaining_time = game_duration - elapsed_time

        if current_time - last_frame_time < 0.05:
            root.after(10, run_game)
            return
        last_frame_time = current_time

        frame_width = frame.shape[1]

        frame = cv2.flip(frame, 1)
        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        img = Image.fromarray(frame_rgb)
        imgtk = ImageTk.PhotoImage(image=img)
        camera_label.imgtk = imgtk
        camera_label.config(image=imgtk)

        with mp_hands.Hands(max_num_hands=1, min_detection_confidence=0.7, min_tracking_confidence=0.7) as hands, \
             mp_face_mesh.FaceMesh(max_num_faces=1, min_detection_confidence=0.7, min_tracking_confidence=0.7) as face_mesh:

            hand_results = hands.process(frame_rgb)
            face_results = face_mesh.process(frame_rgb)

            hand_landmarks_list = []
            if hand_results.multi_hand_landmarks:
                for hand_landmarks in hand_results.multi_hand_landmarks:
                    hand_landmarks_list = [(landmark.x, landmark.y, landmark.z) for landmark in hand_landmarks.landmark]

            mouth_landmarks = []
            if face_results.multi_face_landmarks:
                for face_landmarks in face_results.multi_face_landmarks:
                    mouth_landmarks = [face_landmarks.landmark[i] for i in range(13, 17)]

            analyze_movement(hand_landmarks_list, mouth_landmarks, frame_width)

        elapsed_time_label.config(text=f"残り時間: {int(remaining_time)} 秒")
        score_label.config(text=f"スコア: {score}")
        feedback_label.config(text=feedback_text)
        check_end_game()
        root.after(10, run_game)

def stop_game():
    global is_running
    is_running = False
    pygame.mixer.music.stop()

def reset_game():
    global is_running, score, previous_hand_position, last_frame_time
    is_running = False
    score = 0
    previous_hand_position = None
    last_frame_time = 0
    feedback_label.config(text="")
    advice_label.config(text="")
    elapsed_time_label.config(text="残り時間: 180 秒")
    score_label.config(text="スコア: 0")
    pygame.mixer.music.stop()

def trim_image(image):
    bg = Image.new(image.mode, image.size, image.getpixel((0, 0)))
    diff = ImageChops.difference(image, bg)
    bbox = diff.getbbox()
    if bbox:
        return image.crop(bbox)
    return image

def draw_text_with_outline(draw, position, text, font, outline_color, fill_color, outline_width):
    x, y = position
    draw.text((x - outline_width, y - outline_width), text, font=font, fill=outline_color)
    draw.text((x + outline_width, y - outline_width), text, font=font, fill=outline_color)
    draw.text((x - outline_width, y + outline_width), text, font=font, fill=outline_color)
    draw.text((x + outline_width, y + outline_width), text, font=font, fill=outline_color)
    draw.text(position, text, font=font, fill=fill_color)

def create_tooth_button_with_outline(canvas, text, image_path, command, x, y, radius=50, font_size=40):
    tooth_image = Image.open(image_path).convert("RGBA")
    tooth_image = trim_image(tooth_image)
    background = Image.new("RGBA", tooth_image.size, (173, 216, 230, 255))
    tooth_image = Image.alpha_composite(background, tooth_image)
    tooth_image = tooth_image.resize((radius*2, radius*2))

    draw = ImageDraw.Draw(tooth_image)
    try:
        font = ImageFont.truetype("meiryo.ttc", font_size)
    except IOError:
        font = ImageFont.load_default()

    bbox = draw.textbbox((0, 0), text, font=font)
    text_width, text_height = bbox[2] - bbox[0], bbox[3] - bbox[1]
    text_position = ((radius*2 - text_width) / 2, (radius*2 - text_height) / 2)
    draw_text_with_outline(draw, text_position, text, font, outline_color="white", fill_color="black", outline_width=2)

    tooth_photo = ImageTk.PhotoImage(tooth_image)
    image_id = canvas.create_image(x, y, image=tooth_photo)
    canvas.tag_bind(image_id, "<Button-1>", command)
    canvas.image_refs.append(tooth_photo)

root = tk.Tk()
root.title("ハミガキング(HAMIGAKING)")
root.configure(bg="lightblue")

logo_image = Image.open("hamigaking_rogo.png")
logo_image = logo_image.resize((300, 100))
logo_photo = ImageTk.PhotoImage(logo_image)

logo_label = Label(root, image=logo_photo, bg="lightblue")
logo_label.pack(pady=10)

camera_label = Label(root, bg="lightblue")
camera_label.pack()

elapsed_time_label = tk.Label(root, text="残り時間: 180 秒", font=("Arial", 24), fg="red", bg="lightblue")
elapsed_time_label.pack(pady=10)

score_label = tk.Label(root, text="スコア: 0", font=("Arial", 14), bg="lightblue")
score_label.pack()

feedback_label = tk.Label(root, text="", font=("Arial", 14), fg="green", bg="lightblue")
feedback_label.pack()

advice_label = tk.Label(root, text="", font=("Arial", 14), fg="blue", bg="lightblue")
advice_label.pack(pady=20)

button_canvas = Canvas(root, width=600, height=200, bg="lightblue")
button_canvas.pack()
button_canvas.image_refs = []

create_tooth_button_with_outline(button_canvas, "スタート", "tooth_button.png", lambda event: start_game(), 100, 100, font_size=25)
create_tooth_button_with_outline(button_canvas, "終了", "tooth_button.png", lambda event: root.quit(), 300, 100, font_size=25)
create_tooth_button_with_outline(button_canvas, "リセット", "tooth_button.png", lambda event: reset_game(), 500, 100, font_size=25)

root.mainloop()

cap.release()
cv2.destroyAllWindows()
