import os
import glob
import re
import pandas as pd
import matplotlib.pyplot as plt

# plot PEM temps
def plot_pem(filename):
    plt.figure(figsize=(9,6))
    df = pd.read_excel(filename, usecols="A,B,D,F,H,J,L,N,P")
    plt.plot(df.iloc[:, 0], df.iloc[:, 1:], label=['TC0', 'TC1', 'TC2', 'TC3', 'TC4', 'TC12', 'TC13', 'TC14']) # TODO: Label these
    plt.xlim(0, df.iloc[-1, 0])
    plt.ylim(20, 40)

    base = os.path.basename(filename)
    wattage_value = base.split('W')[0]
    wattage = f"{wattage_value} Watts"

    orientation = "horizontal" if "horizontal" in filename else "vertical"
    output_dir = f"figures/{orientation}"
    os.makedirs(output_dir, exist_ok=True)

    plt.title(f"Temperature of Power Electronics Module at Different Locations\nP = {wattage}, Orientation: {orientation.capitalize()}", fontsize="xx-large")
    plt.ylabel("Temperature [˚C]", fontsize="xx-large")
    plt.xlabel("Time [s]", fontsize="xx-large")
    plt.xticks(fontsize=12)
    plt.yticks(fontsize=12)
    plt.legend(title="Thermocouples", loc="lower right", ncol=2, title_fontsize="x-large", fontsize=12)
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

    plt.title(f"Temperature of Flow Inlet and Outlet\nP = {wattage}, Orientation: {orientation.capitalize()}", fontsize="xx-large")
    plt.ylabel("Temperature [˚C]", fontsize="xx-large")
    plt.xlabel("Time [s]", fontsize="xx-large")
    plt.xticks(fontsize=12)
    plt.yticks(fontsize=12)
    plt.legend(title="Thermocouples", loc="lower right", title_fontsize="x-large", fontsize=12)
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

        # Plot main line
        main_line, = plt.plot(time, avg_temp, label=f"{power} W")

        # Use the same color for additional lines
        color = main_line.get_color()

        # Mask for time > 200
        mask_after_200 = time > 200
        if mask_after_200.any():
            avg_after_200 = avg_temp[mask_after_200].mean()
            plt.hlines(avg_after_200, xmin=time.iloc[0], xmax=time.iloc[-1],
                    colors=color, linestyles='--', 
                    label=f"Avg: {avg_after_200:.2f}°C")

        # Max temperature line
        max_temp = avg_temp.max()
        plt.hlines(max_temp, xmin=time.iloc[0], xmax=time.iloc[-1],
                colors=color, linestyles=':', 
                label=f"Max: {max_temp:.2f}°C")
        
    orientation = "Horizontal" if "horizontal" in data_dir.lower() else "Vertical"

    plt.autoscale(enable=True, axis='x', tight=True)
    plt.xlim(left=0)
    plt.ylim(20, 40)
    plt.title(f"Average Temperature of PEM Over Time\nOrientation: {orientation}", fontsize="xx-large")
    plt.xlabel("Time [s]", fontsize="xx-large")
    plt.ylabel("Average Temperature [˚C]", fontsize="xx-large")
    plt.xticks(fontsize=12)
    plt.yticks(fontsize=12)
    plt.legend(title="Power Level", fontsize=12, loc='upper right', ncol=4, title_fontsize="x-large")
    plt.grid(True)
    plt.tight_layout()

    output_dir = f"figures/{orientation.lower()}"
    os.makedirs(output_dir, exist_ok=True)
    plt.savefig(f"{output_dir}/Average_Temperature_by_Power_{orientation}.png", bbox_inches='tight')

def plot_avg_temp_by_power_all(data_dir):
    # Define patterns for orientations
    orientations = ['horizontal', 'vertical']
    line_styles = {'horizontal': '-', 'vertical': '--'}  # Different line styles

    plt.figure(figsize=(10, 6))

    for orientation in orientations:
        sub_dir = os.path.join(data_dir, orientation)
        if not os.path.exists(sub_dir):
            print(f"No directory found for {orientation}. Skipping.")
            continue

        files = glob.glob(os.path.join(sub_dir, "*PEM.xlsx"))
        if not files:
            print(f"No Excel files found in the {orientation} directory.")
            continue

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

        for power, file in power_file_pairs:
            df = pd.read_excel(file, usecols="A,B,D,F,H,J,L,N,P")
            time = df.iloc[:, 0]
            temps = df.iloc[:, 1:]
            avg_temp = temps.mean(axis=1)
            label = f"{power} W ({orientation.capitalize()})"
            plt.plot(time, avg_temp, line_styles[orientation], label=label)

    plt.autoscale(enable=True, axis='x', tight=True)
    plt.xlim(left=0)
    plt.title("Average Temperature of PEM Over Time\nComparing Orientations", fontsize="xx-large")
    plt.xlabel("Time [s]")
    plt.ylabel("Average Temperature [˚C]")
    plt.legend(title="Power Level & Orientation")
    plt.grid(True)
    plt.tight_layout()

    os.makedirs("figures/comparison", exist_ok=True)
    plt.savefig("figures/comparison/Average_Temperature_by_Power_Comparison.png", bbox_inches='tight')

plot_pem("data/horizontal/19.8W_PEM.xlsx")
plot_pem("data/horizontal/97.5W_PEM.xlsx")
plot_pem("data/horizontal/148.66W_PEM.xlsx")
plot_pem("data/horizontal/198.2W_PEM.xlsx")

plot_flowports("data/horizontal/19.8W_IO.xlsx")
plot_flowports("data/horizontal/97.5W_IO.xlsx")
plot_flowports("data/horizontal/148.66W_IO.xlsx")
plot_flowports("data/horizontal/198.2W_IO.xlsx")

plot_pem("data/vertical/20.88W_Vert_PEM.xlsx")
plot_pem("data/vertical/100.16W_Vert_PEM.xlsx")
plot_pem("data/vertical/149.6W_Vert_PEM.xlsx")
plot_pem("data/vertical/198.39W_Vert_PEM.xlsx")

plot_flowports("data/vertical/20.88W_Vert_IO.xlsx")
plot_flowports("data/vertical/100.16W_Vert_IO.xlsx")
plot_flowports("data/vertical/149.6W_Vert_IO.xlsx")
plot_flowports("data/vertical/198.39W_Vert_IO.xlsx")

plot_avg_temp_by_power("data/horizontal")
plot_avg_temp_by_power("data/vertical")

plot_avg_temp_by_power_all("data")
plt.show()