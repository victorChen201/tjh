import matplotlib.pyplot as PLT
from matplotlib.offsetbox import AnnotationBbox, OffsetImage
from matplotlib._png import read_png
PLT.rcParams['savefig.dpi'] = 300 #图片像素
PLT.rcParams['figure.dpi'] = 300 #分辨率
PLT.figure(figsize=(20,20))
fig = PLT.gcf()
fig.clf()
ax = PLT.subplot(111)

# add a first image
arr_hand = read_png('logo_20200611212020.png')
imagebox = OffsetImage(arr_hand, zoom=.1)
print(arr_hand.shape)

xy = [0.5, 0.5]               # coordinates to position this image

ab = AnnotationBbox(imagebox, xy,
    xybox=(0., 0),
    xycoords='data',
    boxcoords="offset points")
ax.add_artist(ab)

# add second image
arr_vic = read_png('logo_20200611212020.png')
imagebox = OffsetImage(arr_vic, zoom=.1)
# xy = [.6, .3]                  # coordinates to position 2nd image

ab = AnnotationBbox(imagebox, xy,
    xybox=(-10, -10),
    xycoords='data',
    boxcoords="offset points")
print(fig.dpi)
ax.add_artist(ab)

# rest is just standard matplotlib boilerplate
ax.grid(True)
PLT.draw()
PLT.show()