/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
using DocumentFormat.OpenXml.Packaging;

namespace HtmlToOpenXml.IO;

/// <summary>
/// Represents an image and its metadata.
/// </summary>
sealed class HtmlImageInfo(string partId)
{
    /// <summary>
    /// The Unique identifier of the ImagePart in the <see cref="MainDocumentPart"/>.
    /// </summary>
    public string ImagePartId { get; set; } = partId;

    /// <summary>
    /// Gets or sets the original size of the image.
    /// </summary>
    public Size Size { get; set; }

    /// <summary>
    /// Gets the content type of the image.
    /// </summary>
    public PartTypeInfo TypeInfo { get; set; }

    /// <summary>
    /// Gets or sets whether this image is linked externally rather than embedded.
    /// When true, <see cref="ImagePartId"/> contains an external relationship ID instead of an embedded image part ID.
    /// </summary>
    public bool IsExternal { get; set; }
}
