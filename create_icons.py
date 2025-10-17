#!/usr/bin/env python3
"""
Generate neon-style hashtag icons for HashFix Outlook add-in
Creates icons in multiple sizes with a glowing neon effect
"""

from PIL import Image, ImageDraw, ImageFont, ImageFilter
import os

def create_neon_hashtag_icon(size, filename):
    """
    Create a neon-style hashtag icon

    Args:
        size: Icon size in pixels (will be square)
        filename: Output filename
    """
    # Create image with transparent background
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))

    # Create a separate layer for the glow effect
    glow = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw_glow = ImageDraw.Draw(glow)

    # Create main drawing layer
    draw = ImageDraw.Draw(img)

    # Neon color - bright cyan/electric blue
    neon_color = (0, 255, 255, 255)  # Cyan
    glow_color = (0, 200, 255, 180)  # Slightly dimmer cyan for glow

    # Calculate proportions based on size
    padding = max(2, size // 8)
    line_width = max(2, size // 10)

    # Draw hashtag using rectangles
    # Vertical bars
    v1_x = padding + (size - 2 * padding) // 3
    v2_x = padding + 2 * (size - 2 * padding) // 3
    v_width = line_width

    # Horizontal bars
    h1_y = padding + (size - 2 * padding) // 3
    h2_y = padding + 2 * (size - 2 * padding) // 3
    h_height = line_width

    # Draw glow layer (larger, blurred)
    glow_padding = max(1, line_width // 2)

    # Vertical bars for glow
    draw_glow.rectangle(
        [v1_x - glow_padding, padding, v1_x + v_width + glow_padding, size - padding],
        fill=glow_color
    )
    draw_glow.rectangle(
        [v2_x - glow_padding, padding, v2_x + v_width + glow_padding, size - padding],
        fill=glow_color
    )

    # Horizontal bars for glow
    draw_glow.rectangle(
        [padding, h1_y - glow_padding, size - padding, h1_y + h_height + glow_padding],
        fill=glow_color
    )
    draw_glow.rectangle(
        [padding, h2_y - glow_padding, size - padding, h2_y + h_height + glow_padding],
        fill=glow_color
    )

    # Apply blur to glow layer
    blur_radius = max(1, size // 16)
    glow = glow.filter(ImageFilter.GaussianBlur(radius=blur_radius))

    # Composite glow onto main image
    img = Image.alpha_composite(img, glow)

    # Draw sharp hashtag on top
    draw = ImageDraw.Draw(img)

    # Vertical bars
    draw.rectangle(
        [v1_x, padding, v1_x + v_width, size - padding],
        fill=neon_color
    )
    draw.rectangle(
        [v2_x, padding, v2_x + v_width, size - padding],
        fill=neon_color
    )

    # Horizontal bars
    draw.rectangle(
        [padding, h1_y, size - padding, h1_y + h_height],
        fill=neon_color
    )
    draw.rectangle(
        [padding, h2_y, size - padding, h2_y + h_height],
        fill=neon_color
    )

    # Add outer glow/halo effect
    for i in range(3):
        temp_glow = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        temp_draw = ImageDraw.Draw(temp_glow)
        alpha = 40 - (i * 10)

        # Draw slightly larger version for halo
        offset = i + 1
        temp_draw.rectangle(
            [v1_x - offset, padding - offset, v1_x + v_width + offset, size - padding + offset],
            fill=(0, 255, 255, alpha)
        )
        temp_draw.rectangle(
            [v2_x - offset, padding - offset, v2_x + v_width + offset, size - padding + offset],
            fill=(0, 255, 255, alpha)
        )
        temp_draw.rectangle(
            [padding - offset, h1_y - offset, size - padding + offset, h1_y + h_height + offset],
            fill=(0, 255, 255, alpha)
        )
        temp_draw.rectangle(
            [padding - offset, h2_y - offset, size - padding + offset, h2_y + h_height + offset],
            fill=(0, 255, 255, alpha)
        )

        temp_glow = temp_glow.filter(ImageFilter.GaussianBlur(radius=1))
        img = Image.alpha_composite(img, temp_glow)

    # Save the icon
    img.save(filename, 'PNG')
    print(f"Created {filename} ({size}x{size})")

def main():
    """Generate all icon sizes"""
    script_dir = os.path.dirname(os.path.abspath(__file__))

    sizes = [
        (16, 'icon-16.png'),
        (32, 'icon-32.png'),
        (64, 'icon-64.png'),
        (80, 'icon-80.png'),
    ]

    print("Generating neon hashtag icons for HashFix...")
    print("=" * 50)

    for size, filename in sizes:
        filepath = os.path.join(script_dir, filename)
        create_neon_hashtag_icon(size, filepath)

    print("=" * 50)
    print("All icons created successfully!")
    print("\nIcon specifications:")
    print("- Style: Neon glow effect")
    print("- Color: Electric cyan (#00FFFF)")
    print("- Symbol: Hashtag (#)")
    print("- Background: Transparent")

if __name__ == '__main__':
    main()
