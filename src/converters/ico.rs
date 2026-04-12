//! Image to ICO converter module
//! 
//! This module provides functionality to convert various image formats (PNG, JPEG, BMP, WEBP)
//! to ICO (Windows Icon) format with multiple resolutions.

use crate::error::{ConverterError, ConverterResult};
use image::{DynamicImage, ImageFormat};
use ico::{IconDir, IconDirEntry, ResourceDir};
use std::path::Path;

/// Supported input image formats for ICO conversion
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InputImageFormat {
    Png,
    Jpeg,
    Bmp,
    WebP,
    Auto,
}

impl InputImageFormat {
    /// Detect image format from file extension
    pub fn from_path<P: AsRef<Path>>(path: P) -> Option<Self> {
        let ext = path.as_ref().extension()?.to_str()?.to_lowercase();
        match ext.as_str() {
            "png" => Some(InputImageFormat::Png),
            "jpg" | "jpeg" => Some(InputImageFormat::Jpeg),
            "bmp" => Some(InputImageFormat::Bmp),
            "webp" => Some(InputImageFormat::WebP),
            _ => None,
        }
    }

    /// Convert to image::ImageFormat
    pub fn to_image_format(self) -> Option<ImageFormat> {
        match self {
            InputImageFormat::Png => Some(ImageFormat::Png),
            InputImageFormat::Jpeg => Some(ImageFormat::Jpeg),
            InputImageFormat::Bmp => Some(ImageFormat::Bmp),
            InputImageFormat::WebP => Some(ImageFormat::WebP),
            InputImageFormat::Auto => None,
        }
    }
}

/// Configuration for ICO conversion
#[derive(Debug, Clone)]
pub struct IcoConfig {
    /// Generate multiple resolutions (16x16, 32x32, 48x48, 64x64, 128x128, 256x256)
    pub multi_resolution: bool,
    /// Specific sizes to generate (overrides multi_resolution if provided)
    pub sizes: Vec<u16>,
    /// Compression quality (not typically used for ICO, but kept for consistency)
    pub quality: u8,
}

impl Default for IcoConfig {
    fn default() -> Self {
        Self {
            multi_resolution: true,
            sizes: vec![],
            quality: 90,
        }
    }
}

impl IcoConfig {
    /// Create a new IcoConfig with default settings
    pub fn new() -> Self {
        Self::default()
    }

    /// Enable multi-resolution icon generation
    pub fn with_multi_resolution(mut self, enabled: bool) -> Self {
        self.multi_resolution = enabled;
        self
    }

    /// Set specific sizes for the icon
    pub fn with_sizes(mut self, sizes: Vec<u16>) -> Self {
        self.sizes = sizes;
        self.multi_resolution = false;
        self
    }

    /// Get the list of sizes to generate
    pub fn get_sizes(&self) -> Vec<u16> {
        if !self.sizes.is_empty() {
            self.sizes.clone()
        } else if self.multi_resolution {
            vec![16, 32, 48, 64, 128, 256]
        } else {
            vec![256] // Default single size
        }
    }
}

/// Convert an image file to ICO format
/// 
/// # Arguments
/// * `input_path` - Path to the input image file (PNG, JPEG, BMP, WEBP)
/// * `output_path` - Path to the output ICO file
/// 
/// # Returns
/// * `Ok(())` on success
/// * `Err(ConverterError)` on failure
/// 
/// # Example
/// ```no_run
/// use std::path::Path;
/// use file_converter::convert_image_to_ico;
/// 
/// convert_image_to_ico(Path::new("logo.png"), Path::new("logo.ico"))?;
/// ```
pub fn convert_image_to_ico<P: AsRef<Path>, Q: AsRef<Path>>(
    input_path: P,
    output_path: Q,
) -> ConverterResult<()> {
    let config = IcoConfig::default();
    convert_image_to_ico_with_config(input_path, output_path, &config)
}

/// Convert an image file to ICO format with custom configuration
/// 
/// # Arguments
/// * `input_path` - Path to the input image file
/// * `output_path` - Path to the output ICO file
/// * `config` - Conversion configuration
/// 
/// # Returns
/// * `Ok(())` on success
/// * `Err(ConverterError)` on failure
pub fn convert_image_to_ico_with_config<P: AsRef<Path>, Q: AsRef<Path>>(
    input_path: P,
    output_path: Q,
    config: &IcoConfig,
) -> ConverterResult<()> {
    let input_path = input_path.as_ref();
    let output_path = output_path.as_ref();

    // Validate input file exists
    if !input_path.exists() {
        return Err(ConverterError::FileNotFound(input_path.to_path_buf()));
    }

    // Detect or determine input format
    let input_format = if config.quality == 0 {
        InputImageFormat::Auto
    } else {
        InputImageFormat::from_path(input_path).unwrap_or(InputImageFormat::Auto)
    };

    // Load the image
    let img = if input_format == InputImageFormat::Auto {
        image::open(input_path).map_err(|e| {
            ConverterError::ConversionFailed(format!("Failed to open image: {}", e))
        })?
    } else {
        let fmt = input_format.to_image_format().ok_or_else(|| {
            ConverterError::ConversionFailed("Unsupported input format".to_string())
        })?;
        image::load_from_memory_with_format(
            &std::fs::read(input_path).map_err(|e| {
                ConverterError::IoError(std::io::Error::new(
                    std::io::ErrorKind::Other,
                    format!("Failed to read file: {}", e),
                ))
            })?,
            fmt,
        ).map_err(|e| {
            ConverterError::ConversionFailed(format!("Failed to decode image: {}", e))
        })?
    };

    // Get target sizes
    let sizes = config.get_sizes();
    
    if sizes.is_empty() {
        return Err(ConverterError::ConversionFailed(
            "No sizes specified for ICO generation".to_string(),
        ));
    }

    // Create icon directory entries
    let mut entries: Vec<IconDirEntry> = Vec::new();

    for &size in &sizes {
        // Resize image to target size
        let resized = img.resize_exact(size as u32, size as u32, image::imageops::FilterType::Lanczos3);
        
        // Convert to RGBA
        let rgba = resized.to_rgba8();
        
        // Create icon entry
        let entry = IconDirEntry::from_rgba(
            size,
            size,
            &rgba,
        ).map_err(|e| {
            ConverterError::ConversionFailed(format!("Failed to create icon entry: {}", e))
        })?;
        
        entries.push(entry);
    }

    // Create icon directory
    let icon_dir = IconDir::new(entries.len() as u16);
    
    // Write ICO file
    let mut file = std::fs::File::create(output_path).map_err(|e| {
        ConverterError::IoError(std::io::Error::new(
            std::io::ErrorKind::Other,
            format!("Failed to create output file: {}", e),
        ))
    })?;

    // Write icon directory
    icon_dir.write(&mut file).map_err(|e| {
        ConverterError::ConversionFailed(format!("Failed to write icon directory: {}", e))
    })?;

    // Write icon entries
    for entry in &entries {
        entry.write(&mut file).map_err(|e| {
            ConverterError::ConversionFailed(format!("Failed to write icon entry: {}", e))
        })?;
    }

    Ok(())
}

/// Convert image to ICO with a specific size
/// 
/// # Arguments
/// * `input_path` - Path to the input image file
/// * `output_path` - Path to the output ICO file
/// * `size` - Icon size in pixels (e.g., 16, 32, 64, 128, 256)
/// 
/// # Returns
/// * `Ok(())` on success
/// * `Err(ConverterError)` on failure
pub fn convert_image_to_ico_size<P: AsRef<Path>, Q: AsRef<Path>>(
    input_path: P,
    output_path: Q,
    size: u16,
) -> ConverterResult<()> {
    let config = IcoConfig::default().with_sizes(vec![size]);
    convert_image_to_ico_with_config(input_path, output_path, &config)
}

/// Convert image to multi-resolution ICO
/// 
/// Generates an ICO file with multiple resolutions (16x16, 32x32, 48x48, 64x64, 128x128, 256x256)
/// suitable for various Windows UI elements.
/// 
/// # Arguments
/// * `input_path` - Path to the input image file
/// * `output_path` - Path to the output ICO file
/// 
/// # Returns
/// * `Ok(())` on success
/// * `Err(ConverterError)` on failure
pub fn convert_image_to_ico_multi<P: AsRef<Path>, Q: AsRef<Path>>(
    input_path: P,
    output_path: Q,
) -> ConverterResult<()> {
    let config = IcoConfig::default().with_multi_resolution(true);
    convert_image_to_ico_with_config(input_path, output_path, &config)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    #[test]
    fn test_input_image_format_detection() {
        assert_eq!(InputImageFormat::from_path("test.png"), Some(InputImageFormat::Png));
        assert_eq!(InputImageFormat::from_path("test.jpg"), Some(InputImageFormat::Jpeg));
        assert_eq!(InputImageFormat::from_path("test.jpeg"), Some(InputImageFormat::Jpeg));
        assert_eq!(InputImageFormat::from_path("test.bmp"), Some(InputImageFormat::Bmp));
        assert_eq!(InputImageFormat::from_path("test.webp"), Some(InputImageFormat::WebP));
        assert_eq!(InputImageFormat::from_path("test.txt"), None);
    }

    #[test]
    fn test_ico_config_default() {
        let config = IcoConfig::default();
        assert!(config.multi_resolution);
        assert!(config.sizes.is_empty());
        assert_eq!(config.quality, 90);
        assert_eq!(config.get_sizes(), vec![16, 32, 48, 64, 128, 256]);
    }

    #[test]
    fn test_ico_config_custom_sizes() {
        let config = IcoConfig::default().with_sizes(vec![32, 64]);
        assert!(!config.multi_resolution);
        assert_eq!(config.get_sizes(), vec![32, 64]);
    }

    #[test]
    fn test_ico_config_single_size() {
        let config = IcoConfig::default().with_multi_resolution(false);
        assert_eq!(config.get_sizes(), vec![256]);
    }
}
