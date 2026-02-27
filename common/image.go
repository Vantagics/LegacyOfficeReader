package common

// ImageFormat 表示图片的格式类型
type ImageFormat int

const (
	ImageFormatEMF  ImageFormat = iota // Enhanced Metafile
	ImageFormatWMF                     // Windows Metafile
	ImageFormatPICT                    // Macintosh PICT
	ImageFormatJPEG                    // JPEG
	ImageFormatPNG                     // PNG
	ImageFormatDIB                     // Device Independent Bitmap
	ImageFormatTIFF                    // TIFF
)

// Image 表示从文档中提取的单张图片
type Image struct {
	Format ImageFormat // 图片格式
	Data   []byte      // 图片原始字节数据
}

// Extension 返回该图片格式对应的建议文件扩展名
func (img *Image) Extension() string {
	switch img.Format {
	case ImageFormatEMF:
		return ".emf"
	case ImageFormatWMF:
		return ".wmf"
	case ImageFormatPICT:
		return ".pict"
	case ImageFormatJPEG:
		return ".jpeg"
	case ImageFormatPNG:
		return ".png"
	case ImageFormatDIB:
		return ".bmp"
	case ImageFormatTIFF:
		return ".tiff"
	default:
		return ""
	}
}
