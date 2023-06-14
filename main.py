from trc import TRCData
import pandas as pd
from fpdf import FPDF
import subprocess
import math
from datetime import datetime
import argparse


BODY_PARTS = ['sagBilek', 'sagDirsek', 'sagOmuz', 'sagKopr',
              'solBilek', 'solDirsek', 'solOmuz', 'solKopr',
              'omurga1', 'omurga2', 'omurga3', 'omurga4', 'omurga5']

BOX_PARTS = ['side1', 'side2', 'top']

DESK = ['sagOn']

# Vicon uses z-up

def load_trc(filepath: str):
    mocap_data = TRCData()
    mocap_data.load(filepath)

    return mocap_data

def load_xlsx(filepath: str):
    mocap_data = pd.read_excel(filepath, header=[0, 1])
    mocap_data = mocap_data.fillna(mocap_data.mean())

    return mocap_data

def distance_from_body(data):
    distances = []

    for frame in list(data['Frame#']['Frame#']):
        x_body, y_body, z_body = 0, 0, 0
        x_box, y_box, z_box = 0, 0, 0

        for part in BODY_PARTS:
            x_body += data[part]['X'][frame - 1]
            y_body += data[part]['Y'][frame - 1]
            z_body += data[part]['Z'][frame - 1]

        x_body /= len(BODY_PARTS)
        y_body /= len(BODY_PARTS)
        z_body /= len(BODY_PARTS)

        for part in BOX_PARTS:
            x_box += data[part]['X'][frame - 1]
            y_box += data[part]['Y'][frame - 1]
            z_box += data[part]['Z'][frame - 1]

        x_box /= len(BOX_PARTS)
        y_box /= len(BOX_PARTS)
        z_box /= len(BOX_PARTS)

        distance = abs(x_box - x_body) + abs(y_box - y_body) + abs(z_box - z_body)
        distances.append(distance)

    return distances

def calculate_rwl(distances, height):
    rwls = []
    for distance in distances:
        rwl = 23.58 * (distance / height) ** 0.25
        rwls.append(rwl)
    return rwls

def calculate_li(weight, rwls):
    lis = []
    for rwl in rwls:
        li = weight / rwl
        lis.append(li)
    return lis

def calculate_niosh_lifting(weight, distances, height):
    rwls = calculate_rwl(distances, height)
    lis = calculate_li(weight, rwls)
    rwls_dict = {i + 1: rwl for i, rwl in enumerate(rwls)}
    lis_dict = {i + 1: li for i, li in enumerate(lis)}
    return rwls_dict, lis_dict




def interpret_li_values(lis):
    interpretation = []
    errors = []

    for index, li in lis.items():
        try:
            if li < 1.0:
                interpretation.append(f"For index {index}: LI is less than 1.0 - Task is acceptable.")
            elif li == 1.0:
                interpretation.append(f"For index {index}: LI is equal to 1.0 - Task is at the threshold level.")
            else:
                interpretation.append(f"For index {index}: LI is greater than 1.0 - Task exceeds recommended limits.")

        except KeyError:
            errors.append(f"LI value not found for index {index}.")

    return interpretation, errors

def create_report(frames, actions, filename, total_frames, individual_height, lift_weight):
    pdf = FPDF()

    # Set up the PDF document
    pdf.add_page()

    # Add logo image at the top left
    pdf.image('logo.png', x=15, y=15, w=30)

    # Set font and style for the heading
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 20, 'Lifting Analysis Report', ln=True, align='C')

    # Add date under the heading
    pdf.set_font('Arial', '', 12)
    pdf.cell(0, 10, datetime.now().strftime('%Y-%m-%d %H:%M:%S'), ln=True, align='C')

    # Add spacing
    pdf.ln(10)

    # Set font and style for individual and lift weight information
    pdf.set_font('Arial', '', 12)

    # Add individual height information
    pdf.cell(0, 10, '', ln=True, align='L')  # Add spacing
    pdf.cell(0, 10, f'Individual Height: {individual_height} cm', ln=True, align='L')

    # Add lifted weight information
    pdf.cell(0, 5, '', ln=True, align='L')  # Add spacing
    pdf.cell(0, 10, f'Lifted Weight: {lift_weight} kg', ln=True, align='L')
    pdf.ln(10)

    # Set font and style for the table
    pdf.set_font('Arial', 'B', 10)

    # Create table header
    header_height = 10
    pdf.cell(20, header_height, 'Second', 1, 0, 'C')
    pdf.cell(90, header_height, 'Action', 1, 0, 'C')
    pdf.cell(35, header_height, 'Error Frames', 1, 0, 'C')
    pdf.cell(35, header_height, '% Error Frames', 1, 1, 'C')

    # Set font for table content
    pdf.set_font('Arial', '', 10)

    # Calculate the approximate height based on font size
    content_height = max(pdf.font_size, pdf.font_size, pdf.font_size)

    # Keep track of unique seconds and frames with errors
    unique_seconds = set()
    frames_with_errors = set(frames)

    # Calculate the percentage of frames with errors for each second
    error_percentages = []
    for second in range(1, int(math.ceil(total_frames / 120)) + 1):
        frames_in_second = range((second - 1) * 120 + 1, second * 120 + 1)
        error_frames_in_second = frames_with_errors.intersection(frames_in_second)
        error_percentage = len(error_frames_in_second) / 120 * 100
        error_percentages.append(error_percentage)

    # Add frame details, recommended actions, and error percentages
    for frame in frames:
        second = math.ceil((frame - 1) / 120)  # Convert frame to seconds and round up
        action = actions[frame]

        # Calculate the number of error frames and error percentage for the second
        frames_in_second = range((second - 1) * 120 + 1, second * 120 + 1)
        error_frames = len(frames_with_errors.intersection(frames_in_second))
        error_percentage = error_percentages[second - 1]

        # Check if the second is already added
        if second in unique_seconds:
            continue

        # Add the second to the set of unique seconds
        unique_seconds.add(second)

        pdf.cell(20, content_height, str(second), 1, 0, 'C')
        pdf.cell(90, content_height, action, 1, 0, 'L')
        pdf.cell(35, content_height, str(error_frames), 1, 0, 'C')
        pdf.cell(35, content_height, f'{error_percentage:.2f}%', 1, 1, 'C')

    # Adjust table position to the center of the page with margin from the logo
    table_height = len(unique_seconds) * content_height
    top_margin = (pdf.h - table_height) / 2 - 15  # Adjust top margin based on table height
    pdf.set_y(top_margin)

    # Save the PDF file
    pdf.output(filename)

    print(f"Report generated successfully as {filename}")

def suggest_actions_to_maintain_li(weight, distances, height, rwls, lis):
    actions = {}

    for index, li in lis.items():
        if li >= 1.0:
            # Check if weight reduction is feasible
            if weight >= rwls[index]:
                # Reduce weight or use lifting aid
                actions[index] = "Reduce weight or use lifting aid"

                # Check if increasing distance is feasible
                distance_threshold = math.ceil(height * (rwls[index] / 23.58) * 1.2)

                if distances[index - 1] < distance_threshold:
                    # Increase distance and reduce weight or use lifting aid
                    actions[index] += " or increase distance"
                else:
                    # Reduce weight or use lifting aid
                    pass
        else:
            actions[index] = "No action required"

    return actions


def main(filepath, individual_height, lift_weight):
    mocap_data = load_xlsx(filepath)

    distances = distance_from_body(mocap_data)
    
    omurga5z = list(mocap_data['omurga5']['Z'])[0]
    
    cm_mult = omurga5z / 2
    height_cm = ((list(mocap_data['top']['Z'])[-1] - list(mocap_data['top']['Z'])[0]) *  cm_mult) / omurga5z
    rwls, lis = calculate_niosh_lifting(weight=lift_weight, distances=distances, height=height_cm)
    
    interpretation_lis, errors_lis = interpret_li_values(lis)

    for frame, li, interpretation_l in zip(list(mocap_data['Frame#']['Frame#']), list(lis.values()), interpretation_lis):
        print(f"At frame {frame}, for weight {lift_weight}kg, LI:{li}, Interpretation: {interpretation_l}")

    # Create report for frames with RWL and LI out of limits
    frames_out_of_limit = [frame for frame, rwl in rwls.items() if rwl >= 51 or lis[frame] > 1.0]
    actions = suggest_actions_to_maintain_li(lift_weight, distances, individual_height, rwls, lis)
    current_datetime = datetime.now().strftime('%Y%m%d%H%M%S')
    filename = f"lifting_analysis_{current_datetime}.pdf"
    total_frames = len(list(mocap_data['Frame#']['Frame#']))
    create_report(frames_out_of_limit, actions, filename, total_frames, individual_height, lift_weight)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Lifting Analysis Script')
    parser.add_argument('--filepath', type=str, help='Path to the input file')
    parser.add_argument('--height', type=int, help='Individual height in cm')
    parser.add_argument('--weight', type=int, help='Weight in kg')
    args = parser.parse_args()

    filepath = args.filepath
    individual_height = args.height
    lift_weight = args.weight

    main(filepath, individual_height, lift_weight)