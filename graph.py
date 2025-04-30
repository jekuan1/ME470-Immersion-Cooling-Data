import pandas as pd
import matplotlib.pyplot as plt
import os

# plot PEM temps
def plot_pem(filename):
    plt.figure()
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

    plt.title(f"Temperature of Power Electronics Module at Different Locations\nP = {wattage}")
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

    plt.title(f"Temperature of Flow Inlet and Outlet\nP = {wattage}")
    plt.ylabel("Temperature [˚C]")
    plt.xlabel("Time [s]")
    plt.legend()
    plt.grid(True)
    plt.tight_layout()

    output_filename = f"{output_dir}/{wattage} Flowports.png"
    plt.savefig(output_filename)

plot_pem("data/horizontal/19.8W_PEM.xlsx")
plot_pem("data/horizontal/97.5W_PEM.xlsx")
plot_pem("data/horizontal/148.66W_PEM.xlsx")
plot_pem("data/horizontal/198.2W_PEM.xlsx")

plot_flowports("data/horizontal/19.8W_IO.xlsx")
plot_flowports("data/horizontal/97.5W_IO.xlsx")
plot_flowports("data/horizontal/148.66W_IO.xlsx")
plot_flowports("data/horizontal/198.2W_IO.xlsx")

plot_pem("data/vertical/20W_Vert_PEM.xlsx")
plot_pem("data/vertical/100W_Vert_PEM.xlsx")
plot_pem("data/vertical/150W_Vert_PEM.xlsx")
plot_pem("data/vertical/200W_Vert_PEM.xlsx")

plot_flowports("data/vertical/20W_Vert_IO.xlsx")
plot_flowports("data/vertical/100W_Vert_IO.xlsx")
plot_flowports("data/vertical/150W_Vert_IO.xlsx")
plot_flowports("data/vertical/200W_Vert_IO.xlsx")
plt.show()