#!/usr/bin/env python3

from datetime import datetime, timedelta
from PIL import Image, ImageOps
import argparse
import re
import srt
import time
import win32print
import struct

MAXWIDTH = 190
WIDTHCORRECTION = 1.5

ESC = b'\x1B'
GS = b'\x1D'
RESET = ESC + b'@'
SETBOLD = ESC + b'E\x01'
EXITBOLD = ESC + b'E\x00'
CENTER = ESC + b'a\x01'
PAGEFEED = ESC + b'd\x05'
CHARSIZE = GS + b'!'

# 16 pixel line spacing
LS16 = ESC + b'3\x10'

parser = argparse.ArgumentParser(description='Prints text on POS printers.')
parser.add_argument('printer', help='printer name')
parser.add_argument('script', help='script file', type=argparse.FileType('r', encoding='utf-8'))
args = parser.parse_args()

subs = args.script.read()

# Trim UTF-8 BOM if present, or the SRT parser chokes
if len(subs) > 0 and subs[0] == '\ufeff':
	subs = subs[1:]

subs = list(srt.parse(subs))

# Just making sure it's not fucked up
subs.sort(key=lambda x: x.start)

class Line:
	def __init__(self, time, data):
		self.time = time
		self.data = data

	def __repr__(self):
		return 'Line(time=%s, data=%s)' % (self.time, self.data)

startTime = datetime.now()
lines = list()

for sub in subs:
	isFirst = True
	for line in sub.content.split('\n'):
		line = line.lstrip('\r')

		imageMatch = re.match(r'[ \t]*\[img=(.*)\][ \t]*$', line)
		if imageMatch:
			image = Image.open(imageMatch.group(1))

			# Convert into grayscale if not already so we can invert it
			image = image.convert('L')

			# Rescale to fix aspect ratio and make it fit
			if image.width * WIDTHCORRECTION > MAXWIDTH:
				correctheight = round((image.height * MAXWIDTH) / (image.width * WIDTHCORRECTION))
				image = image.resize((MAXWIDTH, correctheight))
			else:
				image = image.resize((round(image.width * WIDTHCORRECTION), image.height))

			# Invert now, as in ESC/POS a 1 is black and 0 is white
			image = ImageOps.invert(image)

			# Create a new black and white image
			bwimage = Image.new('1', (MAXWIDTH, image.height or 7), 0)

			# Paste image centered
			pastepos = (
					round(bwimage.width / 2.0 - image.width / 2.0),
					round(bwimage.height / 2.0 - image.height / 2.0)
			)
			bwimage.paste(image, pastepos)

			# Rotate for slicing
			bwimage = bwimage.transpose(Image.ROTATE_270)
			bwimage = bwimage.transpose(Image.FLIP_LEFT_RIGHT)

			isFirst = True
			header = ESC + b'*\x00' + struct.pack('<H', MAXWIDTH)
			for rowStart in range(0, bwimage.width, 8):
				rowimage = bwimage.crop((rowStart, 0, rowStart + 8, bwimage.height))
				rowdata = bytearray()
				if isFirst:
					# 16 pixels of line spacing (8 pixels of image due to half resolution)
					rowdata.extend(RESET + LS16)
					isFirst = False
				rowdata.extend(header)
				rowdata.extend(rowimage.tobytes())
				rowdata.extend(b'\n')

				lines.append(Line(sub.start, rowdata))
		elif line == '[pagefeed]':
			for i in range(8):
				lines.append(Line(sub.start, b'\n'))
		else:

			if line == '_':
				line = b''
			else:
				line = line.encode('ascii')
				line = line.replace(b'[pagefeed]', PAGEFEED)
				line = line.replace(b'[center]', CENTER)
				line = line.replace(b'[b]', SETBOLD)
				line = line.replace(b'[/b]', EXITBOLD)

				# This is to account for big text that span more than one line
				dummylines = 0

				def sizeCallback(match):
					global dummylines
					size = int(match.group(1)) - 1
					dummylines = max(dummylines, size)
					size = size << 4 | size
					return CHARSIZE + bytes([size])
				line = re.sub(br'\[size=([1-8])\]', sizeCallback, line)

				for x in range(dummylines):
					print('Adding dummy %d' % x)
					lines.append(Line(sub.start, b''))
			if isFirst:
				line = RESET + LS16 + line
				isFirst = False
			lines.append(Line(sub.start, line + b'\n'))

print(lines)

# First "n" lines aren't immediately visible, we have to print ahead of time
timebuffer = [timedelta()] * 3
for line in lines:
	timebuffer.append(line.time)
	line.time = timebuffer.pop(0)

for timestamp in timebuffer:
	lines.append(Line(timestamp, b'\n'))

# Merge lines with common times, so we don't have to open and close the printer so often
curline = lines[0]
mergedlines = list()
for line in lines[1:]:
	if line.time == curline.time:
		curline.data += line.data
	else:
		mergedlines.append(curline)
		curline = line
mergedlines.append(curline)

print(mergedlines)

if True:
	startTime = datetime.now()
	p = win32print.OpenPrinter(args.printer)
	for line in mergedlines:

		delay = line.time + startTime - datetime.now()
		if delay.days >= 0:
			time.sleep(delay.total_seconds())

		win32print.StartDocPrinter(p, 1, ('Line document', None, 'raw'))
		print(line)
		win32print.WritePrinter(p, line.data)
		#win32print.FlushPrinter(p, bytes(line.data), 0)
		win32print.EndDocPrinter(p)

	win32print.ClosePrinter(p)

