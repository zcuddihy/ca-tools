#%%
import pandas as pd
import numpy as np
from calculations import (
    zone_tie_length,
    zone_bar_length,
    vert_distributed_bar_length,
    horiz_distributed_bar_length,
)


class WallQTO:
    def __init__(self, excel_file="./wall_data.xlsx", clrCover=0.040):
        self.levels = pd.read_excel(excel_file, sheet_name="levels").dropna()
        self.walls = (
            pd.read_excel(excel_file, sheet_name="wallProps")
            .drop(columns={"wall", "level"})
            .dropna()
        )
        self.zones = (
            pd.read_excel(excel_file, sheet_name="zoneBar")
            .drop(columns="Unnamed: 0")
            .dropna()
        )
        self.verts = pd.read_excel(excel_file, sheet_name="verticalBar").dropna()
        self.horiz = pd.read_excel(excel_file, sheet_name="horizontalBar").dropna()
        # self.headers = pd.read_excel(excel_file, sheet_name="headers")
        self.barProps = pd.read_excel(excel_file, sheet_name="barProps")
        self.clrCover = clrCover
        self.rebar_wt500 = 0
        self.rebar_wt400 = 0
        self.conc_vol = 0
        self.results = {}

        # Convert levels, wall props and bar props to dictionary
        self.levels.set_index("level", inplace=True)
        self.levels = self.levels.to_dict()
        self.barProps.set_index("bar", inplace=True)
        self.barProps = self.barProps.to_dict()
        self.walls.set_index("index", inplace=True)
        self.wall_list = list(self.walls.index)
        self.walls = self.walls.to_dict()

    def zone_qto(self):

        # Get properties to calculate total wt of vertical zone bar
        # Only half of the bars are spliced at each level
        self.zones["bar_length"] = self.zones.apply(
            lambda row: zone_bar_length(row, self.levels, self.barProps), axis=1
        )
        self.zones["barwt"] = round(
            self.zones["bar_length"]
            * self.zones["bar_size"].map(self.barProps["mass"]),
            0,
        )

        # Compute the total length of ties per story
        self.zones["tie_length"] = self.zones.apply(
            lambda row: zone_tie_length(row, self.walls, self.levels), axis=1
        )

        self.zones["tiewt"] = round(
            self.zones["tie_size"].map(self.barProps["mass"])
            * self.zones["tie_length"],
            0,
        )

        # Combine individual zones into group based on wall/level
        zone_wt = self.zones.groupby("index")["barwt"].sum()
        zone_wt = zone_wt.to_dict()
        tie_wt = self.zones.groupby("index")["tiewt"].sum()
        tie_wt = tie_wt.to_dict()
        return zone_wt, tie_wt

    def vert_distributed_qto(self):
        # Get zone lengths
        self.zones["total_zone_length"] = (
            self.zones["zone_quantity"] * self.zones["zone_length"]
        )
        zone_lengths = self.zones.groupby("index")["total_zone_length"].sum()
        zone_lengths = zone_lengths.to_dict()

        # Compute the total length of ties per story
        self.verts["bar_length"] = self.verts.apply(
            lambda row: vert_distributed_bar_length(
                row, self.walls, self.levels, self.barProps, zone_lengths
            ),
            axis=1,
        )

        self.verts["barwt"] = round(
            self.verts["bar_length"]
            * self.verts["bar_size"].map(self.barProps["mass"]),
            0,
        )

        # Convert to dictionary
        self.verts.set_index("index", inplace=True)
        verts_wt = self.verts["barwt"].to_dict()

        return verts_wt

    def horiz_distributed_qto(self):

        # Calculate total bar length (with a lap splice included) and the total bar weight
        self.horiz["bar_length"] = self.horiz.apply(
            lambda row: horiz_distributed_bar_length(
                row, self.walls, self.levels, self.barProps
            ),
            axis=1,
        )
        self.horiz["bar_wt"] = round(
            self.horiz["bar_size"].map(self.barProps["mass"])
            * self.horiz["bar_length"],
            0,
        )

        # Convert to dictionary
        self.horiz.set_index("index", inplace=True)
        horiz_wt = self.horiz["bar_wt"].to_dict()

        return horiz_wt

    @staticmethod
    def header_alpha(length, depth, nRows, barDia):
        """Calculates the angle of the header diagonals in radians"""
        hDiagonal = (nRows - 1) * 0.1 + 2 * barDia
        alpha = np.arctan((depth - 2 * 0.04) / length) - np.arcsin(hDiagonal / length)
        return alpha

    @staticmethod
    def header_tie_length(nRows, nCols):
        vertLegs = np.where(nCols <= 3, 2, nCols)
        horizLegs = np.where(nRows <= 3, 2, nRows)
        return 0.1 * (nRows - 1) * vertLegs + 0.1 * (nCols - 1) * horizLegs

    def header_quantity(self):

        # Get the angle of the header diagonal reinforcing
        self.headers["alpha"] = WallQTO.header_alpha(
            self.headers["Length"].values,
            self.headers["Depth"].values,
            self.headers["nRows"].values,
            self.headers["bar_size"].map(self.barProps["Dia"]).values,
        )

        # Calculate the total weight of the diagonal reinforcing bars
        self.headers["bar_lengths"] = 2 * (
            (self.headers["Length"] / 2) / np.cos(self.headers["alpha"])
            + self.headers["bar_size"].map(self.barProps["vert_splice"])
        )

        self.headers["barCount"] = 2 * (self.headers["nRows"] * self.headers["nCols"])

        self.headers["diag_wt"] = self.headers["numHeaders"] * (
            self.headers["barCount"]
            * self.headers["bar_lengths"]
            * self.headers["bar_size"].map(self.barProps["mass"])
        )

        # Calculate the total weight of the ties
        self.headers["tieCount"] = (
            2
            * (
                self.headers["bar_lengths"]
                - 2 * self.headers["bar_size"].map(self.barProps["vert_splice"])
            )
            / 0.1
        )

        self.headers["tie_length"] = WallQTO.header_tie_length(
            self.headers["nRows"], self.headers["nCols"]
        )
        self.headers["tie_wt"] = self.headers["numHeaders"] * (
            self.headers["tie_length"] * self.headers["tieCount"] * 0.785
        )

        # Calculate the total concrete volume
        self.headers["concVol"] = self.headers["numHeaders"] * (
            self.headers["Length"] * (self.headers["Depth"] - 0.3)
        )

        # Convert to dictionary
        header_wt = self.headers.groupby("index")["diag_wt"].sum()
        header_tie_wt = self.headers.groupby("index")["tie_wt"].sum()
        header_conc_vol = self.headers.groupby("index")["concVol"].sum()

        return header_wt, header_tie_wt, header_conc_vol

    def calc_rebar_density(self):
        vert_wt = self.vert_distributed_qto()
        zone_wt, zone_tie_wt = self.zone_qto()
        horiz_wt = self.horiz_distributed_qto()
        # header_wt, header_tie_wt, header_conc_vol = self.header_quantity()

        for wall in self.wall_list:
            rebar500 = zone_wt[wall] if wall in zone_wt else 0
            # rebar500 += header_wt[wall] if wall in header_wt else 0

            rebar400 = vert_wt[wall] + horiz_wt[wall] if wall in vert_wt else 0
            rebar400 += zone_tie_wt[wall] if wall in zone_tie_wt else 0
            # rebar400 += header_tie_wt[wall] if wall in header_tie_wt else 0

            # Concrete volume in spreadsheet is in mm^3
            # Need to convert to m^3
            concVol = self.walls["concrete_volume"][wall] / (1000**3)
            # concVol += header_conc_vol[wall] if wall in header_conc_vol else 0

            self.rebar_wt500 += rebar500
            self.rebar_wt400 += rebar400
            self.conc_vol += concVol

            self.results[wall] = {
                "Rebar500": rebar500,
                "Rebar400": rebar400,
                "ConcVol": concVol,
            }

        rebar_density = round((self.rebar_wt500 + self.rebar_wt400) / self.conc_vol, 0)

        # Convert results to dataframe
        results_df = pd.DataFrame.from_dict(self.results, orient="index")
        results_df["Wall"] = results_df.index.str.split("-").str[0]
        results_df["level"] = results_df.index.str.split("-").str[1]
        results_df.reset_index(inplace=True)
        results_df.drop(columns={"index"}, inplace=True)

        return rebar_density, results_df

    def rebar_weights(self):
        return (
            round(self.rebar_wt500, 0),
            round(self.rebar_wt400, 0),
            round(self.conc_vol, 0),
        )


quantities = WallQTO()


rebar_density, results_df = quantities.calc_rebar_density()
rebar500, rebar400, conc_vol = quantities.rebar_weights()

ties = pd.DataFrame(quantities.zones.groupby(["wall", "zone"])["tiewt"].sum())


ties.to_csv("./zone_tie_quantities.csv")

# %%
rebar500, rebar400, conc_vol = quantities.rebar_weights()
# %%
