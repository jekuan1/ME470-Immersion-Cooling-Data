import os
import glob
import re
import pandas as pd
import matplotlib.pyplot as plt

# plot PEM temps
def plot_pem(filename):
    plt.figure(figsize=(9,6))
    df = pd.read_excel(filename, usecols="A,B,D,F,H,J,L,N,P")
    plt.plot(df.iloc[:, 0], df.iloc[:, 1:], label=['TC0', 'TC1', 'TC2', 'TC3', 'TC4', 'TC12', 'TC13', 'TC14'])
    plt.xlim(0, df.iloc[-1, 0])
    plt.ylim(20, 40)

    base = os.path.basename(filename)
    wattage_value = base.split('W')[0]
    wattage = f"{wattage_value} Watts"

    orientation = "horizontal" if "horizontal" in filename else "vertical"
    output_dir = f"figures/{orientation}"
    os.makedirs(output_dir, exist_ok=True)

    plt.title(f"Temperature of Power Electronics Module at Different Locations\nP = {wattage}, Orientation: {orientation.capitalize()}")
    plt.ylabel("Temperature [˚C]")
    plt.xlabel("Time [s]")
    plt.legend()
    plt.grid(True)
    plt.tight_layout()

    output_filename = f"{output_dir}/{wattage} PEM.png"
    plt.savefig(output_filename)

# plot input output flow temp
def plot_flowports(filename):
    plt.figure()
    df = pd.read_excel(filename, usecols="A,B,D")
    plt.plot(df.iloc[:, 0], df.iloc[:, 1:], label=['Inlet', 'Outlet'])
    plt.xlim(0, df.iloc[-1, 0])
    plt.ylim(20, 25)

    base = os.path.basename(filename)
    wattage_value = base.split('W')[0]
    wattage = f"{wattage_value} Watts"

    orientation = "horizontal" if "horizontal" in filename else "vertical"
    output_dir = f"figures/{orientation}"
    os.makedirs(output_dir, exist_ok=True)

    plt.title(f"Temperature of Flow Inlet and Outlet\nP = {wattage}, Orientation: {orientation.capitalize()}")
    plt.ylabel("Temperature [˚C]")
    plt.xlabel("Time [s]")
    plt.legend()
    plt.grid(True)
    plt.tight_layout()

    output_filename = f"{output_dir}/{wattage} Flowports.png"
    plt.savefig(output_filename)

def plot_avg_temp_by_power(data_dir):
    # Find all matching Excel files
    files = glob.glob(os.path.join(data_dir, "*PEM.xlsx"))
    if not files:
        print("No Excel files found in the directory.")
        return

    # Extract (numeric_power, file_path) tuples
    power_file_pairs = []
    for file in files:
        base = os.path.basename(file)
        match = re.match(r"(\d+(?:\.\d+)?)[Ww]", base)
        if match:
            power = float(match.group(1))
            power_file_pairs.append((power, file))

    # Sort by power level
    power_file_pairs.sort()

    plt.figure(figsize=(10, 6))

    for power, file in power_file_pairs:
        df = pd.read_excel(file, usecols="A,B,D,F,H,J,L,N,P")
        time = df.iloc[:, 0]
        temps = df.iloc[:, 1:]
        avg_temp = temps.mean(axis=1)
        plt.plot(time, avg_temp, label=f"{power} W")

    orientation = "Horizontal" if "horizontal" in data_dir.lower() else "Vertical"

    plt.title(f"Average Temperature of PEM Over Time\nOrientation: {orientation}")
    plt.xlabel("Time [s]")
    plt.ylabel("Average Temperature [˚C]")
    plt.legend(title="Power Level")
    plt.grid(True)
    plt.tight_layout()

    output_dir = f"figures/{orientation.lower()}"
    os.makedirs(output_dir, exist_ok=True)
    plt.savefig(f"{output_dir}/Average_Temperature_by_Power_{orientation}.png", bbox_inches='tight')

# plot_pem("data/horizontal/19.8W_PEM.xlsx")
# plot_pem("data/horizontal/97.5W_PEM.xlsx")
# plot_pem("data/horizontal/148.66W_PEM.xlsx")
# plot_pem("data/horizontal/198.2W_PEM.xlsx")

# plot_flowports("data/horizontal/19.8W_IO.xlsx")
# plot_flowports("data/horizontal/97.5W_IO.xlsx")
# plot_flowports("data/horizontal/148.66W_IO.xlsx")
# plot_flowports("data/horizontal/198.2W_IO.xlsx")

# plot_pem("data/vertical/20W_Vert_PEM.xlsx")
# plot_pem("data/vertical/100W_Vert_PEM.xlsx")
# plot_pem("data/vertical/150W_Vert_PEM.xlsx")
# plot_pem("data/vertical/200W_Vert_PEM.xlsx")

# plot_flowports("data/vertical/20W_Vert_IO.xlsx")
# plot_flowports("data/vertical/100W_Vert_IO.xlsx")
# plot_flowports("data/vertical/150W_Vert_IO.xlsx")
# plot_flowports("data/vertical/200W_Vert_IO.xlsx")

plot_avg_temp_by_power("data/horizontal")
plot_avg_temp_by_power("data/vertical")
plt.show()