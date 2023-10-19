
import cv2
import face_recognition
import numpy as np
import os
import xlwt
from xlwt import Workbook
from datetime import date
import xlrd, xlwt
from xlutils.copy import copy as xl_copy

CurrentFolder = os.getcwd()
image = CurrentFolder+'\\data\\var.png'
image2 = CurrentFolder+'\\data\\bsk.png'
image3 = CurrentFolder+'\\data\\tagu.png'
image4 = CurrentFolder+'\\data\\naveesh.png'
video_capture = cv2.VideoCapture(0)


person1_name = "Person1"
person1_image = face_recognition.load_image_file(image)
person1_face_encoding = face_recognition.face_encodings(person1_image)[0]


person2_name = "Person2"
person2_image = face_recognition.load_image_file(image2)
person2_face_encoding = face_recognition.face_encodings(person2_image)[0]

person3_name = "Person3"
person3_image = face_recognition.load_image_file(image3)
person3_face_encoding = face_recognition.face_encodings(person3_image)[0]

person4_name = "Person4"
person4_image = face_recognition.load_image_file(image4)
person4_face_encoding = face_recognition.face_encodings(person4_image)[0]

known_face_encodings = [
    person1_face_encoding,
    person2_face_encoding,
     person3_face_encoding,
     person4_face_encoding
]
known_face_names = [
    person1_name,
    person2_name,
    person3_name,
    person4_name
]

# Initialize some variables
face_locations = []
face_encodings = []
face_names = []
process_this_frame = True

rb = xlrd.open_workbook('attendence_excel.xls', formatting_info=True) 
wb = xl_copy(rb)
inp = input('Please give current subject lecture name')
sheet1 = wb.add_sheet(inp)
sheet1.write(0, 0, 'Name/Date')
sheet1.write(0, 1, str(date.today()))
row=1
col=0
already_attendence_taken = ""
while True:
            # Grab a single frame of video
            ret, frame = video_capture.read()

            # Resize frame of video to 1/4 size for faster face recognition processing
            small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)

            # Convert the image from BGR color (which OpenCV uses) to RGB color (which face_recognition uses)
            rgb_small_frame = small_frame[:, :, ::-1]

            # Only process every other frame of video to save time
            if process_this_frame:
                # Find all the faces and face encodings in the current frame of video
                face_locations = face_recognition.face_locations(rgb_small_frame)
                face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)

                face_names = []
                for face_encoding in face_encodings:
                    # See if the face is a match for the known face(s)
                    matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
                    name = "Unknown"


                    face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)
                    best_match_index = np.argmin(face_distances)
                    if matches[best_match_index]:
                        name = known_face_names[best_match_index]

                    face_names.append(name)
                    if((already_attendence_taken != name) and (name != "Unknown")):
                     sheet1.write(row, col, name )
                     col =col+1
                     sheet1.write(row, col, "Present" )
                     row = row+1
                     col = 0
                     print("attendence taken")
                     wb.save('attendence_excel.xls')
                     already_attendence_taken = name
                    else:
                     print("next student")
                        
            process_this_frame = not process_this_frame



            for (top, right, bottom, left), name in zip(face_locations, face_names):
                # Scale back up face locations since the frame we detected in was scaled to 1/4 size
                top *= 4
                right *= 4
                bottom *= 4
                left *= 4

                # Draw a box around the face
                cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2)

                # Draw a label with a name below the face
                cv2.rectangle(frame, (left, bottom - 35), (right, bottom), (0, 0, 255), cv2.FILLED)
                font = cv2.FONT_HERSHEY_DUPLEX
                cv2.putText(frame, name, (left + 6, bottom - 6), font, 1.0, (255, 255, 255), 1)


            cv2.imshow('Video', frame)


            if cv2.waitKey(1) & 0xff==ord('q'):
                print("data save")
                break

video_capture.release()
cv2.destroyAllWindows()
