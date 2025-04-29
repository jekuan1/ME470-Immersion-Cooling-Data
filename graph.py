import pandas as pd
import matplotlib.pyplot as plt

# plot PEM temps
def plot_pem(filename):
    plt.figure()
    df = pd.read_excel(filename, usecols="A,B,D,F,H,J,L,N,P")
    # for col in df.columns[1:]:
    #     plt.plot(df.iloc[:, 0], df[col], label=col)

    plt.plot(df.iloc[:, 0], df.iloc[:, 1:], label=['TC0', 'TC1', 'TC2', 'TC3', 'TC4', 'TC12', 'TC13', 'TC14'])
    plt.xlim(0, df.iloc[-1, 0])
    plt.ylim(20, 40)

    wattage = filename.split("W")[0] + " Watts"
    plt.title(f"Temperature of Power Electronics Module at Different Locations\nP = {wattage}")
    plt.ylabel("Temperature [˚C]")
    plt.xlabel("Time [s]")
    plt.legend()
    plt.grid(True)
    plt.tight_layout()
    plt.savefig("figures/" + filename[:-5] + ".png")

# plot input output flow temp
def plot_flowports(filename):
    plt.figure()
    df = pd.read_excel(filename, usecols="A,B,D")
    plt.plot(df.iloc[:, 0], df.iloc[:, 1:], label=['Inlet', 'Outlet'])
    plt.xlim(0, df.iloc[-1, 0])
    plt.ylim(20, 25)

    wattage = filename.split("W")[0] + " Watts"
    plt.title(f"Temperature of Flow Inlet and Outlet\nP = {wattage}")
    plt.ylabel("Temperature [˚C]")
    plt.xlabel("Time [s]")
    plt.legend()
    plt.grid(True)
    plt.tight_layout()
    plt.savefig("figures/" + filename[:-5] + ".png")

plot_pem("19.8W_PEM.xlsx")
plot_pem("97.5W_PEM.xlsx")
plot_pem("148.66W_PEM.xlsx")
plot_pem("198.2W_PEM.xlsx")

plot_flowports("19.8W_IO.xlsx")
plot_flowports("97.5W_IO.xlsx")
plot_flowports("148.66W_IO.xlsx")
plot_flowports("198.2W_IO.xlsx")
plt.show()