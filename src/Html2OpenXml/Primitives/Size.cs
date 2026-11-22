/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
namespace HtmlToOpenXml;

/// <summary>
/// Represents a dimension in 2D coordinate space.
/// </summary>
/// <remarks>
/// Initializes a new instance of the <see cref='Size'/> class from
/// the specified dimensions.
/// </remarks>
struct Size(int width, int height) : IEquatable<Size>
{
    /// <summary>
    /// Initializes a new instance of the <see cref='Size'/> class.
    /// </summary>
    public static readonly Size Empty = new();


    /// <summary>
    /// Tests whether this size has zero width and height.
    /// </summary>
    public readonly bool IsEmpty => Width == 0 && Height == 0;

    /// <summary>
    /// Represents the horizontal component of this size.
    /// </summary>
    public int Width { get; set; } = width;

    /// <summary>
    /// Represents the vertical component of this size.
    /// </summary>
    public int Height { get; set; } = height;

    /// <inheritdoc/>
    public bool Equals(Size other)
    {
        return this.Width == other.Width && this.Height == other.Height;
    }

    /// <summary>
    /// Gets a representation of this size.
    /// </summary>
    public override string ToString()
    {
        return string.Format("({0},{1})", Width, Height);
    }

    /// <summary>
    /// Ensure both the Width and Height are positive.
    /// </summary>
    public readonly bool HasInvalidDimension
        => Width <= 0 || Height <= 0;
}
