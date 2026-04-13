//! Image to ICO converter module

use crate::error::{ConverterError, Result};
use ico::{IconDir, IconDirEntry, IconImage, ResourceType};
use image::imageops::FilterType;
use image::RgbaImage;
use std::fs::File;
use std::path::Path;

/// Configuration for ICO conversion
#[derive(Debug, Clone)]
pub struct IcoConfig {
    pub multi_resolution: bool,
    pub sizes: Vec<u16>,
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
    pub fn new() -> Self {
        Self::default()
    }

    pub fn with_multi_resolution(mut self, enabled: bool) -> Self {
        self.multi_resolution = enabled;
        self
    }

    pub fn with_sizes(mut self, sizes: Vec<u16>) -> Self {
        self.sizes = sizes;
        self.multi_resolution = false;
        self
    }

    pub fn get_sizes(&self) -> Vec<u16> {
        if !self.sizes.is_empty() {
            self.sizes.clone()
        } else if self.multi_resolution {
            vec![16, 32, 48, 64, 128, 256]
        } else {
            vec![256]
        }
    }
}

fn is_supported_image(path: &Path) -> bool {
    match path.extension().and_then(|e| e.to_str()).map(|e| e.to_ascii_lowercase()) {
        Some(ref e) => matches!(e.as_str(), "png" | "jpg" | "jpeg" | "gif" | "webp" | "bmp"),
        None => false,
    }
}

fn rgba_for_size(img: &image::DynamicImage, size: u32) -> RgbaImage {
    if img.width() == size && img.height() == size {
        img.to_rgba8()
    } else {
        img.resize_exact(size, size, FilterType::Triangle).to_rgba8()
    }
}

pub fn convert_image_to_ico<P: AsRef<Path>, Q: AsRef<Path>>(input_path: P, output_path: Q) -> Result<()> {
    let config = IcoConfig::default();
    convert_image_to_ico_with_config(input_path, output_path, &config)
}

pub fn convert_image_to_ico_with_config<P: AsRef<Path>, Q: AsRef<Path>>(
    input_path: P,
    output_path: Q,
    config: &IcoConfig,
) -> Result<()> {
    let input_path = input_path.as_ref();
    let output_path = output_path.as_ref();

    if !input_path.exists() {
        return Err(ConverterError::FileNotFound(input_path.to_string_lossy().to_string()));
    }
    if !is_supported_image(input_path) {
        return Err(ConverterError::UnsupportedFormat(
            input_path
                .extension()
                .and_then(|e| e.to_str())
                .unwrap_or("")
                .to_string(),
        ));
    }

    let img = image::open(input_path).map_err(|e| ConverterError::ConversionFailed(e.to_string()))?;
    let mut icon_dir = IconDir::new(ResourceType::Icon);

    for size in config.get_sizes() {
        let s = u32::from(size);
        let rgba = rgba_for_size(&img, s);
        let rgba_vec = rgba.into_raw();
        let icon_image = IconImage::from_rgba_data(u32::from(size), u32::from(size), rgba_vec);
        let entry = IconDirEntry::encode(&icon_image).map_err(|e| ConverterError::ConversionFailed(e.to_string()))?;
        icon_dir.add_entry(entry);
    }

    let file = File::create(output_path).map_err(ConverterError::from)?;
    icon_dir
        .write(file)
        .map_err(|e| ConverterError::ConversionFailed(e.to_string()))?;

    log::info!("Wrote ICO to {}", output_path.display());
    Ok(())
}

pub fn convert_image_to_ico_size<P: AsRef<Path>, Q: AsRef<Path>>(input_path: P, output_path: Q, size: u16) -> Result<()> {
    let config = IcoConfig::default().with_sizes(vec![size]);
    convert_image_to_ico_with_config(input_path, output_path, &config)
}

pub fn convert_image_to_ico_multi<P: AsRef<Path>, Q: AsRef<Path>>(input_path: P, output_path: Q) -> Result<()> {
    let config = IcoConfig::default().with_multi_resolution(true);
    convert_image_to_ico_with_config(input_path, output_path, &config)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_ico_config() {
        let config = IcoConfig::default();
        assert!(config.multi_resolution);
        assert_eq!(config.get_sizes(), vec![16, 32, 48, 64, 128, 256]);
    }
}
