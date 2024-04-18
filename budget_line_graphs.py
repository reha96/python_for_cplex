#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Sep  7 23:33:10 2022

@author: reha.tuncer
"""

import matplotlib.pyplot as plt
plt.rcParams['figure.dpi'] = 300
import numpy as np


b1x = np.array([0, 40])
b1y = np.array([13.34, 0])

b3x = np.array([0, 60])
b3y = np.array([30, 0])

b5x = np.array([0, 75])
b5y = np.array([37.5, 0])

b7x = np.array([0, 60])
b7y = np.array([60, 0])

b8x = np.array([0, 100])
b8y = np.array([100, 0])

b9x = np.array([0, 80])
b9y = np.array([80, 0])

b10x = np.array([0, 40])
b10y = np.array([10, 0])

plt.plot(b1x,b1y, label='Budget 1')
plt.plot(b1y,b1x, label='Budget 2')

plt.plot(b3x,b3y, label='Budget 3')
plt.plot(b3y,b3x, label='Budget 4')

plt.plot(b5x,b5y, label='Budget 5')
plt.plot(b5y,b5x, label='Budget 6')

plt.plot(b7x,b7y, label='Budget 7')
# plt.plot(b8x,b8y, label='Budget 8')
# plt.plot(b9x,b9y, label='Budget 9')

plt.plot(b10x,b10y, label='Budget 8')
plt.plot(b10y,b10x, label='Budget 9')

plt.xlabel("Vacation good")
plt.ylabel("Shopping good")

plt.tight_layout() 

plt.legend()
plt.xlim([0, 80])
plt.ylim([0, 80])
plt.show()

# TIME BUDGET

b1x = np.array([0, 720])
b1y = np.array([120, 0])

b2x = np.array([0, 240])
b2y = np.array([120, 0])

b3x = np.array([0, 720])
b3y = np.array([180, 0])

b4x = np.array([0, 360])
b4y = np.array([180, 0])

b5x = np.array([0, 900])
b5y = np.array([225, 0])

b6x = np.array([0, 225])
b6y = np.array([225, 0])

b7x = np.array([0, 180])
b7y = np.array([180, 0])

b8x = np.array([0, 600])
b8y = np.array([300, 0])

b9x = np.array([0, 480])
b9y = np.array([240, 0])

b10x = np.array([0, 450])
b10y = np.array([90, 0])

b11x = np.array([0, 720])
b11y = np.array([240, 0])

plt.plot(b1x,b1y, label='Budget 1')
plt.plot(b2x,b2y, label='Budget 2')

plt.plot(b3x,b3y, label='Budget 3')
plt.plot(b4x,b4y, label='Budget 4')

plt.plot(b5x,b5y, label='Budget 5')
plt.plot(b6x,b6y, '--', label='Budget 6')

plt.plot(b7x,b7y, label='Budget 7')
plt.plot(b8x,b8y, '--', label='Budget 8')
plt.plot(b9x,b9y, label='Budget 9')

plt.plot(b10x,b10y, label='Budget 10')
plt.plot(b11x,b11y, label='Budget 11')

plt.xlabel("Labor (consumption possibility)")
plt.ylabel("Leisure")

plt.tight_layout() 

plt.legend()
plt.xlim([0, 900])
plt.ylim([0, 300])
plt.show()
