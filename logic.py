import math
import pandas as pd
vpdf = pd.read_csv("ViewportsCSV")
fittingdf = pd.read_csv("FittingsCSV")

    # Check if point p is inside parallelogram abcd
    # ax, ay = a
    # bx, by = b
    # cx, cy = c
    # dx, dy = d
    # px, py = p
    
    # # Calculate vectors
    # ab = (bx - ax, by - ay)
    # bc = (cx - bx, cy - by)
    # cd = (dx - cx, dy - cy)
    # da = (ax - dx, ay - dy)
    # ap = (px - ax, py - ay)
    # bp = (px - bx, py - by)
    
    # # Calculate dot products
    # abap = ab[0] * ap[0] + ab[1] * ap[1]
    # bcbp = bc[0] * bp[0] + bc[1] * bp[1]
    # ccdp = cd[0] * (px - cx) + cd[1] * (py - cy)
    # dadp = da[0] * (px - dx) + da[1] * (py - dy)
    
    # # Check if point is inside parallelogram
    # if 0 <= abap <= ab[0]**2 + ab[1]**2 and 0 <= bcbp <= bc[0]**2 + bc[1]**2 \
    # and 0 <= ccdp <= cd[0]**2 + cd[1]**2 and 0 <= dadp <= da[0]**2 + da[1]**2:
        
    #     return True
    # else:
    #     return False
def is_inside_parallelogram(a, b, c, d, p):

    ax, ay = a
    bx, by = b
    cx, cy = c
    dx, dy = d
    px, py = p
    
    # Calculate vectors
    ab = (bx - ax, by - ay)
    ad = (dx - ax, dy - ay)
    ap = (px - ax, py - ay)
    
    # Compute cross products
    ab_ap_cross = ab[0]*ap[1] - ab[1]*ap[0]
    ad_ap_cross = ad[0]*ap[1] - ad[1]*ap[0]
    
    # Check if point is inside parallelogram
    distance = abs((bx-ax)*(ay-py)-(ax-px)*(by-ay)) / math.sqrt((bx-ax)**2 + (by-ay)**2)
    print(distance)
    if (0 <= ab_ap_cross <= ab[0]*ad[1] - ab[1]*ad[0] and
        0 <= ad_ap_cross <= ab[0]*ad[1] - ab[1]*ad[0]):
        return True
    else:
        return False


# unit tests
print(f"w2  {is_inside_parallelogram((2705833.459672579,272065.5911821082),(2706066.2346493243,271629.32848659245), (2706194.218869466,271809.9125254958), (2705705.4754524375,271885.0071432049), (2705755.071673409,271943.5055327221))} ")

print(f"w1  {is_inside_parallelogram((2705391.9902256182,272374.00578466937),(2705684.5322455964,271852.41415261204), (2705833.4596725786,272065.5911821087),(2705243.0627986356,272160.8287551728), (2705755.071673409,271943.5055327221))} ")


def run():   
    # Creating new vpdf Column based on associted Viewport to Fitting
    fittingdf['Matching Viewport ID'] = 'N/A'
    fittingdf['Matching Viewport Sheet'] = 'N/A'

    # Iteratively gather all coordinates by going through Viewports that are 
    # main BasePlan ModelSpaces
    viewportIndex, fittingIndex = 0, 0
    for viewportIndex in vpdf.index:
        if vpdf['Is BasePlan ModelSpace'][viewportIndex] == True:
            # ccw rotation about corners
            # 1 -> 3 -> 2 -> 4
            corner1 = (vpdf['ModelSpace Coordinate Corner1 X'][viewportIndex],
                            vpdf['ModelSpace Coordinate Corner1 Y'][viewportIndex])
            'Sorted in a CW Manner'
            corner2 = (vpdf['ModelSpace Coordinate Corner3 X'][viewportIndex],
                            vpdf['ModelSpace Coordinate Corner3 Y'][viewportIndex])
            corner3 = (vpdf['ModelSpace Coordinate Corner2 X'][viewportIndex],
                            vpdf['ModelSpace Coordinate Corner2 Y'][viewportIndex])
            corner4 = (vpdf['ModelSpace Coordinate Corner4 X'][viewportIndex],
                            vpdf['ModelSpace Coordinate Corner4 Y'][viewportIndex])
            
            for fittingIndex in fittingdf.index:
                fittingPoint = (fittingdf['Block X'][fittingIndex],
                                fittingdf['Block Y'][fittingIndex])
                # Only check for Fittings inside ModelSpace
                if fittingdf['Sheet'][fittingIndex] == 'Model' and fittingdf['Matching Viewport ID'][fittingIndex] == 'N/A':
                    if is_inside_parallelogram(corner1, corner2, corner3, corner4, fittingPoint):
                        fittingdf.loc[fittingIndex, 'Matching Viewport ID'] = \
                            vpdf['ID'][viewportIndex]
                        fittingdf.loc[fittingIndex, 'Matching Viewport Sheet'] = \
                            vpdf['Sheet'][viewportIndex]
                        
fittingdf.to_csv('FittingssSV')
print("--Logged to Fittings")