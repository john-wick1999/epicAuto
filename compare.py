import cv2 as cv
import numpy as np
from matplotlib import pyplot as plt

img = cv.imread('Capture3.jpg', cv.IMREAD_GRAYSCALE)
img2 = img.copy()
template = cv.imread('Screenshot_1.png', cv.IMREAD_GRAYSCALE)
# template = cv.imread('search_bar.JPG', cv.IMREAD_GRAYSCALE)

w, h = template.shape[::-1]
# All the 6 methods for comparison in a list
methods = ['cv.TM_CCOEFF', 'cv.TM_CCOEFF_NORMED', 'cv.TM_CCORR',
            'cv.TM_CCORR_NORMED', 'cv.TM_SQDIFF', 'cv.TM_SQDIFF_NORMED']
results = []
for meth in methods:
    img = img2.copy()
    method = eval(meth)
    # Apply template Matching
    res = cv.matchTemplate(img, template, method)
    min_val, max_val, min_loc, max_loc = cv.minMaxLoc(res)
    # If the method is TM_SQDIFF or TM_SQDIFF_NORMED, take minimum
    
    if method in [cv.TM_SQDIFF, cv.TM_SQDIFF_NORMED]:
        top_left = min_loc
        print("min Trushold {}, {}".format(min_val, max_val > 0.9))
    else:
        print("max Trushold {}. {}".format(max_val, max_val > 0.9))
        
        top_left = max_loc
    bottom_right = (top_left[0] + w, top_left[1] + h)
    
    # Calculate the midpoint between top and bottom
    midpoint = ((top_left[0] + bottom_right[0]) // 2, (top_left[1] + bottom_right[1]) // 2)
    results.append((top_left[1] + bottom_right[1]) // 2)
    
    
    cv.rectangle(img, top_left, bottom_right, 255, 2)
    
    # Mark the midpoint with a red dot
    cv.circle(img, midpoint, 5, (0, 0, 255), -1)
    
    plt.subplot(121), plt.imshow(res, cmap='gray')
    plt.title('Matching Result'), plt.xticks([]), plt.yticks([])
    plt.subplot(122), plt.imshow(img, cmap='gray')
    plt.title(f'Detected Point - Midpoint: {midpoint}'), plt.xticks([]), plt.yticks([])
    plt.suptitle(meth)
    plt.show()

print (results)