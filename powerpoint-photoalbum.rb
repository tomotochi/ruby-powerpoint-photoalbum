require "win32ole"

LONGEST_EDGE = 640
SHORT_EDGE_THRESHOLD = 100
PIXEL_TO_POINT = 0.75
THUMB_PREFIX = "thumb."

#-----------------------------------------------------------------------------
# initialization
#-----------------------------------------------------------------------------

# create OLE instances
ppt = WIN32OLE.new "PowerPoint.Application"
fso = WIN32OLE.new "Scripting.FileSystemObject"
im = WIN32OLE.new "ImageMagickObject.MagickImage.1"

# consts placeholder
class PpConst; end
class MsoConst; end

WIN32OLE.const_load(ppt, PpConst)
WIN32OLE.const_load("Microsoft Office 12.0 Object Library", MsoConst)

#-----------------------------------------------------------------------------
# startup
#-----------------------------------------------------------------------------

# accept first arg as the directory that contains bunch of jpg files
dir = File.expand_path(ARGV.first)
raise unless FileTest.directory?(dir)

files = Dir.glob(File.join(dir, "*.jpg"))
files.reject!{|e| File.basename(e) =~ /^#{THUMB_PREFIX}/}
files.sort!

thumbs = []
orientation_stat = Hash.new(0)

#-----------------------------------------------------------------------------
# convert by ImageMagick
#-----------------------------------------------------------------------------

files.each do |file|
  args = []
  
  # inspect image file
  (width, height) =
    *im.Identify("-format", "%w %h", file).split.map{|e| e.to_i}
  
  # portrait or landscape
  if height > width
    orientation_stat[MsoConst::MsoOrientationVertical] += 1
  else
    orientation_stat[MsoConst::MsoOrientationHorizontal] += 1
  end
  
  # resize if an edge exceeds LONGEST_EDGE, or is too short
  max = [width, height].max
  if max > LONGEST_EDGE || max < SHORT_EDGE_THRESHOLD
    args << "-resize" << "#{LONGEST_EDGE}x#{LONGEST_EDGE}"
    args << "-define" << "jpeg:size=#{LONGEST_EDGE}x#{LONGEST_EDGE}"
  end
  
  # convert to grayscale
  args << "-colorspace" << "Gray"
  
  # thumbnail filename
  thumb = File.join(File.dirname(file), [THUMB_PREFIX, File.basename(file)].join)
  thumbs << thumb
  
  # do ImageMagick convert
  im.Convert(*args, file, thumb)
end

#-----------------------------------------------------------------------------
# build PowerPoint object
#-----------------------------------------------------------------------------

# most frequent orientation
orientation = orientation_stat.sort{|a, b| b.last <=> a.last}.first.first

# create Presentation instance
pre = ppt.Presentations.Add

# setup slide geometry
pre.PageSetup.SlideOrientation = orientation
case orientation
when MsoConst::MsoOrientationHorizontal
  pre.PageSetup.SlideWidth = LONGEST_EDGE * PIXEL_TO_POINT
  pre.PageSetup.SlideHeight = pre.PageSetup.SlideWidth * 0.75
when MsoConst::MsoOrientationVertical
  pre.PageSetup.SlideHeight = LONGEST_EDGE * PIXEL_TO_POINT
  pre.PageSetup.SlideWidth = pre.PageSetup.SlideHeight * 0.75
end

# create "Photo Album"
thumbs.each do |thumb|
  (width, height) =
    *im.Identify("-format", "%w %h", thumb).split.map{|e| e.to_i}
  
  slide = pre.Slides.Add(pre.Slides.Count + 1, PpConst::PpLayoutBlank)
  
  shape = slide.Shapes.AddPicture(
    fso.GetAbsolutePathName(thumb),
    true, false,
    0, 0, width * PIXEL_TO_POINT, height * PIXEL_TO_POINT)
end

# overwrite save
pre.SaveAs(fso.GetAbsolutePathName([dir, "ppt"].join(".")))
pre.Close
ppt.Quit
