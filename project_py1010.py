# The various analyses of data logged for the support department at the telecommunications company MORSE.

import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

# Task A
# A program that reads the file 'support_uke_24.xlsx' and stores the data from columns 1, 2, 3, and 4 in the arrays u_dag, kl_slett, varighet, and score.

# Reads the file
data = pd.read_excel("support_uke_24.xlsx")

# Stores data
u_dag = data["Ukedag"].values
kl_slett = data["Klokkeslett"].values
varighet = data["Varighet"].values
score = data["Tilfredshet"].values

# Check data
# print("Days", u_dag)
# print("Clock", kl_slett)
# print("Duration", varighet)
# print("Score", score)

# Task B
# A program that counts the number of inquiries for each of the five weekdays and visualizes the results using a bar chart.

def find_inquiries(data):
    weekdays = ["Mandag", "Tirsdag", "Onsdag", "Torsdag", "Fredag"]
    
    inquiry_counts = {day: 0 for day in weekdays}
    
    for day in data:
        if day in inquiry_counts:
            inquiry_counts[day] += 1
             
    # Shows result
    plt.figure()
    plt.bar(inquiry_counts.keys(), inquiry_counts.values())   
    plt.xlabel("Ukedag")
    plt.ylabel("Antall henvendelser")
    plt.title("Antall henvendelder per ukedag")
    plt.grid
    plt.savefig("inquiries.pdf", format="pdf")
    plt.show()
    
find_inquiries(u_dag)
    
# Task C
# A program that finds the shortest and longest call duration logged for week 24 and displays the result with informative text.

def find_min_max_duration(data):

    shortest_call = min(data)
    longest_call = max(data)
    
    print("Shortest call: ", shortest_call)
    print("Longest call: ", longest_call)
    
find_min_max_duration(varighet)

# Task D
# A program that calculates the average call duration based on all inquiries in week 24.
 
def find_average_duration(data):
    total_seconds = 0
    count = 0
    
    for time_str in data:
        # Convert string to datetime object
        time_obj = datetime.strptime(time_str, '%H:%M:%S')
        # Convert datetime object to seconds
        seconds = time_obj.hour * 3600 + time_obj.minute * 60 + time_obj.second
        total_seconds += seconds
        count += 1
        
    # Count average in seconds
    average_seconds = total_seconds / count
    
    # Convert to %H:%M:%S
    average_time = f"{int(average_seconds // 3600):02}:{int((average_seconds % 3600) // 60):02}:{int(average_seconds % 60):02}"
       
    print("Count inquiries: ", count)
    print("Average call duration: ", average_time)
    
find_average_duration(varighet)

# Task E
# A program that finds the total number of inquiries received by the support department for each time period (08-10, 10-12, 12-14, and 14-16) in week 24.
# The results are visualized using a pie chart.

def find_inquiries_by_period(data):
    time_intervals = {
        "kl 08-10": 0,
        "kl 10-12": 0,
        "kl 12-14": 0,
        "kl 14-16": 0
    }
    
    for time_str in data:
        # Convert string to datetime object
        time_obj = datetime.strptime(time_str, '%H:%M:%S')
        hour = time_obj.hour
        
        # Count inquiries by period
        if 8 <= hour < 10:
            time_intervals["kl 08-10"] += 1
        elif 10 <= hour < 12:
            time_intervals["kl 10-12"] += 1
        elif 12 <= hour < 14:
            time_intervals["kl 12-14"] += 1
        elif 14 <= hour < 16:
            time_intervals["kl 14-16"] += 1
    
    # check inquiries count
    total = sum(time_intervals.values())
    print("Total inquiries: ", total)
    
    return time_intervals

# show result           
inquiries_by_period = find_inquiries_by_period(kl_slett)
plt.figure(figsize=(7, 7))
plt.pie(
    inquiries_by_period.values(),
    labels=[f"{key}\n{value} ({value/sum(inquiries_by_period.values())*100:.1f}%)" for key, value in inquiries_by_period.items()],
    colors=["#FF9999", "#66B2FF", "#99FF99", "#FFCC99"]
)
plt.title("Fordeling av henvendelser per tidsrom (uke 24)")
plt.savefig("inquiries_by_period.pdf", format="pdf")
plt.show()

# Task F
# A program that calculates the support department's NPS and displays the result on the screen.
# Customers who have not provided feedback on satisfaction will be excluded from the calculations.

def find_nps(data):
    
    nps_clients = {
        "Detractors": 0,
        "Passives": 0,
        "Promoters": 0
    }
    
    total_responses = 0
    
    # Count scores
    for number in data:
        if pd.notna(number): # if not NAN
            total_responses += 1
            if 1 <= number < 7:
                nps_clients["Detractors"] += 1
            elif 7 <= number < 9:
                nps_clients["Passives"] += 1
            elif 9 <= number <= 10:
                nps_clients["Promoters"] += 1
    
    # Count NPS
    
    if total_responses > 0:
        percent_promoters = (nps_clients["Promoters"] / total_responses) * 100
        percent_detractors = (nps_clients["Detractors"] / total_responses) * 100
        percent_passives = (nps_clients["Passives"] / total_responses) * 100
        nps = percent_promoters - percent_detractors
        
        # Show result in pie diagram
        colors = ["#FF6666", "#FFD966", "#66CC99"]
        fig, ax = plt.subplots(figsize=(8, 4))

        ax.pie(
            [percent_detractors, percent_passives, percent_promoters],
            labels=["Detractors", "Passives", "Promoters"],
            autopct="%1.1f%%",
            colors=colors,
            wedgeprops={"edgecolor": "white", "linewidth": 2},
        )
        
        plt.title(f"NPS Score: {int(nps)}", fontsize=14, fontweight="bold")
        plt.savefig("nps_score.pdf", format="pdf")
        plt.show()
    else:
        nps = None
   
    print("Total responses", total_responses)
    
    return nps

nps = find_nps(score)

# Show result
if nps is not None:
    print(f"NPS (Net Promoter Score) is: {nps:.2f}")
else:
    print("No valid satisfaction scores available.")



