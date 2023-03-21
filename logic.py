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

    # cross prod
    # ax, ay = a
    # bx, by = b
    # cx, cy = c
    # dx, dy = d
    # px, py = p
    
    # # Calculate vectors
    # ab = (bx - ax, by - ay)
    # ad = (dx - ax, dy - ay)
    # ap = (px - ax, py - ay)
    
    # # Compute cross products
    # ab_ap_cross = ab[0]*ap[1] - ab[1]*ap[0]
    # ad_ap_cross = ad[0]*ap[1] - ad[1]*ap[0]
    
    # # Check if point is inside parallelogram
    # distance = abs((bx-ax)*(ay-py)-(ax-px)*(by-ay)) / math.sqrt((bx-ax)**2 + (by-ay)**2)
    # print(distance)
    # if (0 <= ab_ap_cross <= ab[0]*ad[1] - ab[1]*ad[0] and
    #     0 <= ad_ap_cross <= ab[0]*ad[1] - ab[1]*ad[0]):
    #     return True
    # else:
    #     return False

def is_inside_quadrilateral(a, b, c, d, p):
    # Compute winding numbers
    wn_abp = compute_winding_number(a, b, p)
    wn_bcp = compute_winding_number(b, c, p)
    wn_cdp = compute_winding_number(c, d, p)
    wn_dap = compute_winding_number(d, a, p)

    # Check if point is inside quadrilateral
    if (wn_abp == wn_bcp == wn_cdp == wn_dap) and (wn_abp != 0):
        return True
    else:
        return False

def compute_winding_number(start, end, point):
    # Compute vector from start to point
    v = (point[0] - start[0], point[1] - start[1])

    # Compute vector from start to end
    w = (end[0] - start[0], end[1] - start[1])

    # Compute the 2D cross product of v and w
    cross_product = (v[0] * w[1]) - (v[1] * w[0])

    # Determine winding number based on sign of cross product
    if cross_product > 0:
        return 1
    elif cross_product < 0:
        return -1
    else:
        return 0
    
def arrange_points_clockwise(a, b, c, d):
    # Find centroid
    centroid_x = (a[0] + b[0] + c[0] + d[0]) / 4
    centroid_y = (a[1] + b[1] + c[1] + d[1]) / 4
    centroid = (centroid_x, centroid_y)

    # Compute angles between centroid and points
    angles = []
    for point in [a, b, c, d]:
        x_diff = point[0] - centroid[0]
        y_diff = point[1] - centroid[1]
        angle = math.atan2(y_diff, x_diff)
        angles.append(angle)

    # Sort points by angle
    sorted_points = [x for _, x in sorted(zip(angles, [a, b, c, d]))]
    return sorted_points

# # False

# a, b, c, d = arrange_points_clockwise((2705391.9902256182,272374.00578466937),(2705684.5322455964,271852.41415261204), (2705833.459672579,272065.5911821082),(2705243.0627986356,272160.8287551728))
# print(is_inside_quadrilateral(a, b, c, d, (2705755.071673409,271943.5055327221)))

# # True
# a, b, c, d = arrange_points_clockwise((2705833.459672579,272065.5911821082),(2706066.2346493243,271629.32848659245), (2706194.218869466,271809.9125254958), (2705705.4754524375,271885.0071432049))
# print(is_inside_quadrilateral(a, b, c, d, (2705755.071673409,271943.5055327221)))




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
                if fittingdf['Sheet'][fittingIndex] == 'Model':
                    #  and fittingdf['Matching Viewport ID'][fittingIndex] == 'N/A'
                    if is_inside_quadrilateral(corner1, corner2, corner3, corner4, fittingPoint):
                        fittingdf.loc[fittingIndex, 'Matching Viewport ID'] = \
                            vpdf['ID'][viewportIndex]
                        fittingdf.loc[fittingIndex, 'Matching Viewport Sheet'] = \
                            vpdf['Sheet'][viewportIndex]
                        

run()                        
fittingdf.to_csv('FittingssSV')
print("--Logged to Fittings")