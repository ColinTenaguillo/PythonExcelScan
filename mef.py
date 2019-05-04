# !/usr/bin/env python
# -*-coding:Utf-8 -*


""".py: Mes Fonctions / class."""

from dataclasses import dataclass


@dataclass
class Service:
    """Class pour les services à analyser."""

    nom: str
    temps: int = 0
    nb: int = 0

    def ajoute(self, nom_ticket, temps):
        """Add value to the variable time."""
        if self.nom in nom_ticket:
            self.temps += temps
            self.nb += 1

    def print_excel(self, i_row, sheet):
        """Add variable to the new sheet."""
        sheet.cell(row=i_row, column=1).value = self.nom
        sheet.cell(row=i_row, column=2).value = self.nb
        sheet.cell(row=i_row, column=3).value = round((self.temps / 60), 2)

    def printall(self):
        """Just a function to print in terminal to see the service value."""
        return (
            self.nom
            + " : "
            + str(self.nb)
            + " : "
            + str(round((self.temps / 60), 2))
        )
