import random
import discord

class Dice_make:
    def dice_make(count):
        dice = random.randint(1, count)
        return dice

    def dice_result(dice1, dice2):
        dice_sum = dice1 + dice2
        if dice_sum <= 6:
            return 0
        else:
            if 6 < dice_sum <= 9:
                return 1
            else:
                return 2