import os
from PIL import Image


def png_to_jpeg(png_path, jpeg_path=None, quality=95):
    """
    将PNG图片转换为JPEG格式
    
    参数:
    png_path (str): PNG图片的路径
    jpeg_path (str, optional): 输出JPEG图片的路径，默认为原路径替换后缀
    quality (int): JPEG图片质量，范围0-100，默认95
    
    返回:
    str: 转换后的JPEG图片路径
    """
    # 检查输入文件是否存在
    if not os.path.exists(png_path):
        raise FileNotFoundError(f"PNG文件不存在: {png_path}")

    # 如果未指定输出路径，则自动生成
    if jpeg_path is None:
        base_name = os.path.splitext(png_path)[0]
        jpeg_path = base_name + ".jpg"

    # 打开PNG图片
    with Image.open(png_path) as img:
        # 如果图片有透明通道，需要转换为RGB模式
        print(f"正在处理{png_path}")
        if img.mode in ('RGBA', 'LA', 'P'):
            # 创建白色背景
            background = Image.new('RGB', img.size, (255, 255, 255))
            # 将PNG图片粘贴到白色背景上
            if img.mode == 'P':
                img = img.convert('RGBA')
            background.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
            img = background
        elif img.mode != 'RGB':
            # 确保图片是RGB模式
            img = img.convert('RGB')

        # 保存为JPEG格式
        img.save(jpeg_path, 'JPEG', quality=quality)

    return jpeg_path


if __name__ == '__main__':
    folder = input("请输入要转换的文件夹路径（直接回车默认为当前目录）：") or "./未命名文件夹"
    if not os.path.exists(folder):
        print(f"文件夹 {folder} 不存在")
        exit(1)
        
    converted_count = 0
    for root, dirs, files in os.walk(folder):
        for file in files:
            if file.lower().endswith(".png"):
                png_path = os.path.join(root, file)
                jpeg_path = os.path.join(root, os.path.splitext(file)[0] + ".jpg")
                try:
                    png_to_jpeg(png_path, jpeg_path)
                    print(f"已转换: {png_path} -> {jpeg_path}")
                    converted_count += 1
                except Exception as e:
                    print(f"转换失败 {png_path}: {e}")
    
    print(f"总共转换了 {converted_count} 个PNG文件")
