"""
Creates a Sasquatch/Bigfoot icon for the Log Analyzer application.
"""

try:
    from PIL import Image, ImageDraw
except ImportError:
    import subprocess
    import sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "Pillow"])
    from PIL import Image, ImageDraw


def create_sasquatch_icon():
    """Create a simple Sasquatch silhouette icon."""

    # Create images at multiple sizes for .ico file
    sizes = [16, 32, 48, 64, 128, 256]
    images = []

    for size in sizes:
        # Create image with transparent background
        img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Scale factor
        s = size / 256

        # Background circle (dark green forest theme)
        padding = int(8 * s)
        draw.ellipse(
            [padding, padding, size - padding, size - padding],
            fill=(34, 85, 51, 255)  # Dark forest green
        )

        # Sasquatch silhouette (dark brown/black)
        sasquatch_color = (30, 20, 15, 255)  # Very dark brown

        # Body proportions scaled to icon size
        center_x = size // 2

        # Head (oval)
        head_width = int(50 * s)
        head_height = int(55 * s)
        head_top = int(35 * s)
        draw.ellipse([
            center_x - head_width // 2,
            head_top,
            center_x + head_width // 2,
            head_top + head_height
        ], fill=sasquatch_color)

        # Body (large oval)
        body_width = int(85 * s)
        body_height = int(100 * s)
        body_top = int(75 * s)
        draw.ellipse([
            center_x - body_width // 2,
            body_top,
            center_x + body_width // 2,
            body_top + body_height
        ], fill=sasquatch_color)

        # Left arm
        arm_points = [
            (center_x - int(40 * s), int(90 * s)),   # Shoulder
            (center_x - int(75 * s), int(130 * s)),  # Elbow out
            (center_x - int(60 * s), int(170 * s)),  # Hand
            (center_x - int(45 * s), int(160 * s)),  # Inner arm
            (center_x - int(35 * s), int(110 * s)),  # Back to body
        ]
        draw.polygon(arm_points, fill=sasquatch_color)

        # Right arm
        arm_points_r = [
            (center_x + int(40 * s), int(90 * s)),
            (center_x + int(75 * s), int(130 * s)),
            (center_x + int(60 * s), int(170 * s)),
            (center_x + int(45 * s), int(160 * s)),
            (center_x + int(35 * s), int(110 * s)),
        ]
        draw.polygon(arm_points_r, fill=sasquatch_color)

        # Left leg
        leg_points = [
            (center_x - int(30 * s), int(155 * s)),  # Hip
            (center_x - int(45 * s), int(200 * s)),  # Knee
            (center_x - int(50 * s), int(235 * s)),  # Foot
            (center_x - int(25 * s), int(235 * s)),  # Foot inner
            (center_x - int(20 * s), int(200 * s)),  # Inner leg
            (center_x - int(15 * s), int(160 * s)),  # Back to body
        ]
        draw.polygon(leg_points, fill=sasquatch_color)

        # Right leg
        leg_points_r = [
            (center_x + int(30 * s), int(155 * s)),
            (center_x + int(45 * s), int(200 * s)),
            (center_x + int(50 * s), int(235 * s)),
            (center_x + int(25 * s), int(235 * s)),
            (center_x + int(20 * s), int(200 * s)),
            (center_x + int(15 * s), int(160 * s)),
        ]
        draw.polygon(leg_points_r, fill=sasquatch_color)

        # Eyes (small glowing dots for mystery)
        eye_color = (200, 180, 100, 255)  # Yellowish glow
        eye_y = int(55 * s)
        eye_size = max(2, int(6 * s))

        # Left eye
        draw.ellipse([
            center_x - int(15 * s) - eye_size // 2,
            eye_y - eye_size // 2,
            center_x - int(15 * s) + eye_size // 2,
            eye_y + eye_size // 2
        ], fill=eye_color)

        # Right eye
        draw.ellipse([
            center_x + int(15 * s) - eye_size // 2,
            eye_y - eye_size // 2,
            center_x + int(15 * s) + eye_size // 2,
            eye_y + eye_size // 2
        ], fill=eye_color)

        images.append(img)

    # Save as .ico file with multiple sizes
    icon_path = "sasquatch.ico"
    images[-1].save(
        icon_path,
        format='ICO',
        sizes=[(s, s) for s in sizes],
        append_images=images[:-1]
    )

    print(f"Icon created: {icon_path}")
    return icon_path


if __name__ == "__main__":
    create_sasquatch_icon()
    print("\nIcon file 'sasquatch.ico' has been created!")
    print("You can now run build_exe.bat to create the executable with this icon.")
