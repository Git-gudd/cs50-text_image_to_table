# import the necessary packages
from sklearn.cluster import AgglomerativeClustering
from pytesseract import Output
from tabulate import tabulate
import pandas as pd
import numpy as np
import pytesseract
import argparse
import imutils
import cv2
import logging

logging.basicConfig(level=logging.INFO)

# construct the argument parser and parse the arguments
ap = argparse.ArgumentParser()
ap.add_argument("-i", "--image", required=True,
	help="path to input image to be OCR'd")
ap.add_argument("-o", "--output", required=True,
	help="path to output CSV file")

# Basic option
ap.add_argument("-r", "--have-header", required=False, type=int, default=0,
	help="if the data has header, set to True")
ap.add_argument("-n", "--remove-character", required=False, type=str, default="",
	help="remove any unwanted specific character")
ap.add_argument("-a", "--column-alignment", required=False, type=int, default=-1,
	help="if the data is align left, set to -1; if align center, set to 0; and if align right, set to 1")

# Other options
ap.add_argument("-c", "--min-conf", required=False, type=int, default=0,
	help="minimum confidence value to filter weak text detection")
ap.add_argument("-f", "--fulltable", required=False, type=int, default=1,
	help="if the image only include the table, then set to True")

# Advance options
ap.add_argument("-fc", "--fixed-col", required=False, type=int, default=0,
	help="specify number of columns in the image")
ap.add_argument("-fr", "--fixed-row", required=False, type=int, default=0,
	help="specify number of rows in the image")
ap.add_argument("-dc", "--col-dist-thresh", required=False, type=float, default=25.0,
	help="distance threshold cutoff for column clustering, use this when no number of columns have been set")
ap.add_argument("-dr", "--row-dist-thresh", required=False, type=float, default=5.0,
	help="distance threshold cutoff for row clustering")
ap.add_argument("-mc", "--min-col-cell", required=False, type=int, default=2,
	help="minimum cluster size (i.e., # of entries in a column)")
ap.add_argument("-mr", "--min-row-cell", required=False, type=int, default=1,
	help="minimum cluster size (i.e., # of entries in a row)")

class NotEnoughCol(Exception):
    "Raised when the min value of number of columns is higher that the text detected"
    # In order words, the text detected does not form enough columns or rows required
    pass

class NotEnoughRow(Exception):
    "Raised when the min value of number of rows is higher that the text detected"
    # In order words, the text detected does not form enough columns or rows required
    pass

def remove_character(data, characters):
	new_data = []
	for text in data:
		new_text = text.replace(characters, "")
		new_data.append(new_text)
	return new_data		

def main(args):
	### Set a seed for our random number generator to generate colors for each column of text we detect 
	np.random.seed(42)


	### Load the input image 
	logging.debug(args["image"].encode("utf-8"))
	original_image = cv2.imread(args["image"])

	# If the image contains other text outside the table
	if args["fulltable"] == 0:
		def crop_table_from_image(image):
			# convert it to grayscale
			gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

			# initialize a rectangular kernel that is ~5x wider than it is tall,
			# then smooth the image using a 3x3 Gaussian blur and then apply a
			# blackhat morphological operator to find dark regions on a light background
			kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (51, 11))
			gray = cv2.GaussianBlur(gray, (3, 3), 0)
			blackhat = cv2.morphologyEx(gray, cv2.MORPH_BLACKHAT, kernel)

			# compute the Scharr gradient of the blackhat image and scale the
			# result into the range [0, 255]
			grad = cv2.Sobel(blackhat, ddepth=cv2.CV_32F, dx=1, dy=0, ksize=-1)
			grad = np.absolute(grad)
			(minVal, maxVal) = (np.min(grad), np.max(grad))
			grad = (grad - minVal) / (maxVal - minVal)
			grad = (grad * 255).astype("uint8")


			# apply a closing operation using the rectangular kernel to close
			# gaps in between characters, apply Otsu's thresholding method, and
			# finally a dilation operation to enlarge foreground regions
			grad = cv2.morphologyEx(grad, cv2.MORPH_CLOSE, kernel)
			thresh = cv2.threshold(grad, 0, 255,
				cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
			thresh = cv2.dilate(thresh, None, iterations=3)

			cv2.imshow("Thresh", thresh)



			# Phase 3
			#!!!!!
			# find contours in the thresholded image and grab the largest one,
			# which we will assume is the stats table
			cnts = cv2.findContours(thresh.copy(), cv2.RETR_EXTERNAL,
				cv2.CHAIN_APPROX_SIMPLE)
			cnts = imutils.grab_contours(cnts)
			tableCnt = max(cnts, key=cv2.contourArea)
			# compute the bounding box coordinates of the stats table and extract
			# the table from the input image
			(x, y, w, h) = cv2.boundingRect(tableCnt)
			table = image[y:y + h, x:x + w]
			# show the original input image and extracted table to our screen
			cv2.imshow("Input", image)
			cv2.imshow("Table", table)

			return table
		table = crop_table_from_image(original_image)
	# If the whole image only contain the table
	else:
		table = original_image
	### Using pyteseract to find text in image

	# set the PSM mode to detect sparse text, and then localize text in the table
	options = "--psm 4"
	tesseract_result = pytesseract.image_to_data(cv2.cvtColor(table, cv2.COLOR_BGR2RGB),
				     							config=options,
												output_type=Output.DICT)
	logging.debug(tesseract_result)

	# initialize a list to store the (x, y)-coordinates of the detected
	# text along with the OCR'd text itself
	def get_coord_n_text(tesseract_result, args):
		coords = []
		ocrText = []

		# loop over each of the individual text localizations
		for i in range(0, len(tesseract_result["text"])):
			# extract the bounding box coordinates of the text region from
			# the current result
			x = tesseract_result["left"][i]
			y = tesseract_result["top"][i]
			w = tesseract_result["width"][i]
			h = tesseract_result["height"][i]
			# extract the OCR text itself along with the confidence of the
			# text localization
			text = tesseract_result["text"][i]
			conf = int(tesseract_result["conf"][i])
			# filter out weak confidence text localizations
			if conf > args["min_conf"]:
				# update our text bounding box coordinates and OCR'd text,
				# respectively
				coords.append((x, y, w, h))
				ocrText.append(text)
			
		return coords, ocrText
	coords, ocrText = get_coord_n_text(tesseract_result, args)

	# Group texts's coordinate close to each other into columns
	def group_x_coords_into_columns(coords, args, image):
		# Group texts's coordinate close to each other into columns

		# If the number of columns have been specify:
		if args["fixed_col"] != 0:
			xclustering = AgglomerativeClustering(
				n_clusters=args["fixed_col"],
				affinity="manhattan",
				linkage="complete",
				distance_threshold=None)
		
		# Or apply hierarchical agglomerative clustering to the coordinates
		else:
			def get_dist_thresh(thresh_power, image_width):
				thresh_base = image_width ** (1/10)
				x_dist_thresh = thresh_base ** thresh_power
				return x_dist_thresh
			
			xclustering = AgglomerativeClustering(
				n_clusters=None,
				affinity="manhattan",
				linkage="complete",
				distance_threshold=get_dist_thresh(args["col_dist_thresh"], image.shape[1]))
			
		def collect_x_coords(coords, alignment):
			### Detect columns in the image
			
			# @ to detect column that align right
			# extract all x-coordinates + width of text boxes
			if alignment == 1:
				xCoords = [(c[0] + c[2], 0) for c in coords]

			# @ to detect column that align center
			# (extract all x-coordinates + width of text boxes) / 2
			elif alignment == 0:
				xCoords = [((c[0] + c[2])/2, 0) for c in coords]

			# @ to detect column that align left - also default behaviour
			# extract all x-coordinates from the text bounding boxes, setting the y-coordinate value to zero
			else:
				xCoords = [(c[0], 0) for c in coords]

			return xCoords
		xCoords = collect_x_coords(coords, args["column_alignment"])
		xclustering.fit(xCoords)

		return xclustering
	def group_y_coords_into_rows(coords, args, image):
		# Group texts's coordinate close to each other into rows
		if args["fixed_row"] != 0:
			yclustering = AgglomerativeClustering(
				n_clusters=args["fixed_row"],
				affinity="manhattan",
				linkage="complete",
				distance_threshold=None)
		else:
			def get_dist_thresh(thresh_power, image_width):
				thresh_base = image_width ** (1/10)
				x_dist_thresh = thresh_base ** thresh_power
				return x_dist_thresh
			
			yclustering = AgglomerativeClustering(
				n_clusters=None,
				affinity="manhattan",
				linkage="complete",
				distance_threshold=get_dist_thresh(args["row_dist_thresh"], image.shape[0]))
		
		yCoords = [(c[1], 0) for c in coords]
		yclustering.fit(yCoords)

		return yclustering
	xclustering = group_x_coords_into_columns(coords, args, table)
	yclustering = group_y_coords_into_rows(coords, args, table)
	

	###_______________Sort and rearrange column/row in the order they appeared_______________
	def sort_x_clustering(xclustering, coords, args):
		# initialize our list of sorted clusters
		sorted_xClusters = []

		# Phase : Looping and sorting clusters
		# loop over all clusters
		for l in np.unique(xclustering.labels_):
			# extract the indexes for the coordinates belonging to the
			# current cluster
			idxs = np.where(xclustering.labels_ == l)[0]
			# verify that the cluster is sufficiently large
			if len(idxs) >= args["min_row_cell"]:
				# compute the average x-coordinate value of the cluster and
				# update our clusters list with the current label and the
				# average x-coordinate
				avg = np.average([coords[i][0] for i in idxs])
				sorted_xClusters.append((l, avg))
		
		if len(sorted_xClusters) == 0:
			raise NotEnoughRow
		
		# sort the clusters by their average x-coordinate and 
		sorted_xClusters.sort(key=lambda x: x[1])
		for c in sorted_xClusters:
			logging.debug(f"X clusterings: {c}")

		return sorted_xClusters
	def sort_y_clustering(yclustering, coords, args):
		# initialize our list of sorted clusters
		sorted_yClusters = []

		# Phase : Looping and sorting clusters
		# loop over all clusters
		for r in np.unique(yclustering.labels_):
			# extract the indexes for the coordinates belonging to the
			# current cluster
			idxs = np.where(yclustering.labels_ == r)[0]
			# verify that the cluster is sufficiently large
			if len(idxs) >= args["min_col_cell"]:
				# compute the average y-coordinate value of the cluster and
				# update our clusters list with the current label and the
				# average x-coordinate
				avg = np.average([coords[i][1] for i in idxs])
				sorted_yClusters.append((r, avg))
				
		if len(sorted_yClusters) == 0:
			raise NotEnoughCol
		
		# sort the clusters by their average x-coordinate
		sorted_yClusters.sort(key=lambda x: x[1])
		for r in sorted_yClusters:
			logging.debug(f"Y clusterings: {r}")

		return sorted_yClusters
	sorted_xClusters = sort_x_clustering(xclustering, coords, args)	
	sorted_yClusters = sort_y_clustering(yclustering, coords, args)
	

	###________________________Initialize dataframe with rows and columns________________________
	# Initialize data frame
	df = pd.DataFrame() 
	
	# Create another list of the text arrange in rows
	def create_sub_y_index(coords):
		yclustering = AgglomerativeClustering(
				n_clusters=None,
				affinity="manhattan",
				linkage="complete",
				distance_threshold=10)
		
		yCoords = [(c[1], 0) for c in coords]
		yclustering.fit(yCoords)

		sorted_yClusters = []
		for r in np.unique(yclustering.labels_):
			idxs = np.where(yclustering.labels_ == r)[0]
			avg = np.average([coords[i][1] for i in idxs])
			sorted_yClusters.append((r, avg))
		sorted_yClusters.sort(key=lambda x: x[1])

		ordered_yIdxs = []
		for col_index, (r, _) in enumerate(sorted_yClusters):
			idxs = np.where(yclustering.labels_ == r)[0]
			ordered_yIdxs.append(idxs)

		return ordered_yIdxs
	sub_yIdxs = create_sub_y_index(coords)

	# Loop over the clusters again, this time in sorted order, so that cells are arrange in the correct column/row order
	# Loop over every columns
	for col_index, (c, _) in enumerate(sorted_xClusters):
		xIdxs = np.where(xclustering.labels_ == c)[0]
		col_data = []
		
		# Generate a random color for the column
		color = np.random.randint(0, 255, size=(3,), dtype="int")
		color = [int(c) for c in color]

		# Loop over every rows
		for (r, _) in sorted_yClusters:
			yIdxs = np.where(yclustering.labels_ == r)[0]

			# Get the indexes of every text in the current cell
			xyIdxs = np.intersect1d(xIdxs, yIdxs)
			if len(xyIdxs) == 0:
				col_data.append('')
				continue

			# Loop over every lines in cell
			def get_sortedIdxs_in_cell(xyIdxs, sub_yIdxs):
				sorted_cellIdxs = []
				for row in sub_yIdxs:

					# Get the indexes of every text in the current line
					lineIdxs = np.intersect1d(xyIdxs, row)
					if len(lineIdxs) == 0:
						continue

					text_xCoords = [coords[i][0] for i in lineIdxs]
					sorted_lineIdxs = lineIdxs[np.argsort(text_xCoords)]
					sorted_cellIdxs.append(sorted_lineIdxs)
				return sorted_cellIdxs
			sorted_cellIdxs = get_sortedIdxs_in_cell(xyIdxs, sub_yIdxs)

			def get_text_from_Idxs(cellIdxs):
				cell_line_data = []
				for lineIdxs in cellIdxs:
					textLine = [ocrText[i].strip() for i in lineIdxs]
					cell_line_data.append(' '.join(textLine))
				
				return '\n'.join(cell_line_data)
			col_data.append(get_text_from_Idxs(sorted_cellIdxs))

			# Draw rectangle on output image
			def get_cell_border(xyIdxs):
				leftCoords = [coords[i][0] for i in xyIdxs]
				topCoords = [coords[i][1] for i in xyIdxs]
				rightCoords = [coords[i][0] + coords[i][2] for i in xyIdxs]
				bottomCoords = [coords[i][1] + coords[i][3] for i in xyIdxs]
				x1 = min(leftCoords)
				y1 = min(topCoords)
				x2 = max(rightCoords)
				y2 = max(bottomCoords)
				return (x1, y1), (x2, y2)
			(x1, y1), (x2, y2) = get_cell_border(xyIdxs)
			cv2.rectangle(original_image, (x1, y1), (x2, y2), color, 2)
				

	###______________________Extract text from each cell to a df__________________________
		# extract the OCR'd text for the current column, then construct
		# a data frame for the data where the first entry in our column
		# serves as the header
		if args["remove_character"]:
			col_data = remove_character(col_data, args["remove_character"])

		if args["have_header"] == 1:
			currentDF = pd.DataFrame({col_data[0]: col_data[1:]})
		else:
			currentDF = pd.DataFrame({col_index: col_data})


		# concatenate *original* data frame with the *current* data
		# frame (we do this to handle columns that may have a varying
		# number of rows)
		df = pd.concat([df, currentDF], axis=1)

	###_____________________________Save data to file______________________________
	def save_data_to_file(df, args):	
		# replace NaN values with an empty string and then show a nicely
		# formatted version of our multi-column OCR'd text
		df.fillna("", inplace=True)

		if args["have_header"] == 1:
			result_table = tabulate(df, headers="keys", tablefmt="psql")
			
		else:
			result_table = tabulate(df, tablefmt="psql")

		logging.info(f"\n{result_table}")
		# write our table to disk as a CSV file
		logging.debug("Saving CSV file to disk...")
		df.to_csv(args["output"], index=False)
	save_data_to_file(df, args)
	
	# Return the output image after performing multi-column OCR
	return cv2.cvtColor(table, cv2.COLOR_BGR2RGB)

if __name__ == "__main__":
	args = vars(ap.parse_args())
	image = main(args)

	cv2.imshow("Output", image)
	cv2.waitKey(0)
