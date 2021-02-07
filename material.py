from apo_interface import Logistical_Material_Data
import math

class Material:
    def __init__(self, fpc, qty):
        self.fpc = fpc
        self.apo_md = Logistical_Material_Data(self.fpc)
        self.description = self.apo_md.product_description
        self.quantity = qty
        self.packaging = "pallet"
        self.pallet_length = 80
        self.pallet_width = 120
        self.pallet_height_P2 = self.apo_md.P2_height
        self.pallet_height_airfreight = self.apo_md.P2_height
        self.reducing_factor = 0
        while ( self.pallet_height_airfreight > 160 ):
            self.pallet_height_airfreight -= self.apo_md.LE_height
            self.reducing_factor += 1
        self.layers_per_af_pallet = math.floor( self.pallet_height_airfreight / self.apo_md.LE_height )
        self.items_per_af_pallet = self.apo_md.LE_items * self.layers_per_af_pallet
        self.weight_per_af_pallet = self.apo_md.P2_weight - (self.reducing_factor * self.apo_md.LE_weight)
        self.pallet_count = qty / self.items_per_af_pallet
        self.weight = self.weight_per_af_pallet * self.pallet_count
