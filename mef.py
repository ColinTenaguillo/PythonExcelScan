#!/usr/bin/env python
# -*-coding:Utf-8 -*


""".py: Mes Fonctions / class."""

from dataclasses import dataclass


@dataclass
class Service:
    """Class to create a service."""

    nom: str
    temps: int = 0
    nb: int = 0

    def ajoute(self, nom_ticket, temps):
        """Add data in temps and nb of the service."""
        if self.nom in nom_ticket:
            self.temps += temps
            self.nb += 1

    def print_excel(self, i_row, sheet):
        """Add data in the excel file."""
        sheet.cell(row=i_row, column=1).value = self.nom
        sheet.cell(row=i_row, column=2).value = self.nb
        sheet.cell(row=i_row, column=3).value = round((self.temps / 60), 2)

    def printall(self):
        """Print all data of a service in bash."""
        return (
            self.nom
            + " : "
            + str(self.nb)
            + " : "
            + str(round((self.temps / 60), 2))
        )
