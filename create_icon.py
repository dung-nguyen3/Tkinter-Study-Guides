#!/usr/bin/env python3
"""
Create a chart icon for the Excel Master Chart Creator app
"""

from PIL import Image, ImageDraw, ImageFont
import os

def create_chart_icon(size=1024):
    """Create a colorful chart/graph icon"""

    # Create image with rounded corners and gradient background
    img = Image.new('RGBA', (size, size), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)

    # Background with gradient effect (medical/study theme colors)
    bg_color = (45, 125, 210)  # Blue

    # Draw rounded rectangle background
    margin = size // 20
    draw.rounded_rectangle(
        [(margin, margin), (size - margin, size - margin)],
        radius=size // 8,
        fill=bg_color
    )

    # Draw a simple bar chart
    chart_margin = size // 5
    chart_bottom = size - chart_margin
    chart_left = chart_margin
    chart_right = size - chart_margin
    chart_top = chart_margin + size // 10

    # Draw bars (representing a chart)
    bar_count = 5
    bar_width = (chart_right - chart_left) // (bar_count * 2)
    bar_spacing = bar_width // 2

    colors = [
        (255, 107, 107),  # Red
        (255, 195, 113),  # Orange
        (106, 176, 76),   # Green
        (79, 172, 254),   # Light Blue
        (162, 155, 254),  # Purple
    ]

    heights = [0.7, 0.5, 0.85, 0.6, 0.75]  # Relative heights

    x = chart_left + bar_spacing
    for i in range(bar_count):
        bar_height = heights[i] * (chart_bottom - chart_top)
        bar_top = chart_bottom - bar_height

        # Draw bar with slight 3D effect
        draw.rectangle(
            [(x, bar_top), (x + bar_width, chart_bottom)],
            fill=colors[i]
        )

        # Add white outline to bars
        draw.rectangle(
            [(x, bar_top), (x + bar_width, chart_bottom)],
            outline=(255, 255, 255, 200),
            width=size // 200
        )

        x += bar_width + bar_spacing

    # Draw axis lines
    axis_color = (255, 255, 255, 230)
    axis_width = size // 150

    # Y-axis
    draw.line(
        [(chart_left, chart_top), (chart_left, chart_bottom + size // 100)],
        fill=axis_color,
        width=axis_width
    )

    # X-axis
    draw.line(
        [(chart_left - size // 100, chart_bottom), (chart_right, chart_bottom)],
        fill=axis_color,
        width=axis_width
    )

    # Add small grid lines for detail
    grid_color = (255, 255, 255, 80)
    for i in range(1, 4):
        y = chart_bottom - (i * (chart_bottom - chart_top) // 4)
        draw.line(
            [(chart_left, y), (chart_right, y)],
            fill=grid_color,
            width=axis_width // 2
        )

    return img

def save_icon_sizes(base_img, output_dir):
    """Save icon in multiple sizes for macOS"""

    sizes = [16, 32, 64, 128, 256, 512, 1024]
    icon_files = []

    for size in sizes:
        resized = base_img.resize((size, size), Image.Resampling.LANCZOS)
        filename = os.path.join(output_dir, f'icon_{size}x{size}.png')
        resized.save(filename, 'PNG')
        icon_files.append(filename)
        print(f"Created {filename}")

    return icon_files

if __name__ == '__main__':
    print("Creating chart icon...")

    # Create the base icon
    icon = create_chart_icon(1024)

    # Save main icon
    icon.save('app_icon.png', 'PNG')
    print("Created app_icon.png")

    # Save multiple sizes for iconset
    os.makedirs('AppIcon.iconset', exist_ok=True)
    icon_files = save_icon_sizes(icon, 'AppIcon.iconset')

    print("\nIcon created successfully!")
    print("To create .icns file, run:")
    print("  iconutil -c icns AppIcon.iconset -o app_icon.icns")
