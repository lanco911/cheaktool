def calculate_apportionment(building_areas, common_areas):
    """
    计算分摊面积和分摊系数。

    参数:
    building_areas (dict): 一个字典，键是房屋门牌号，值是每户的建筑面积。
    common_areas (dict): 一个字典，键是公共建筑区域的名称，值是它们对应的建筑面积。

    返回:
    tuple: 一个分摊后的楼层面积字典和分摊系数。
    """
    # 计算总建筑面积
    total_building_area = sum(building_areas.values())

    # 计算总公共区域面积
    total_common_area = sum(common_areas.values())

    # 计算分摊系数
    apportionment_factor = round(total_common_area / total_building_area, 7)

    # 为每户计算分摊后的楼层面积，并保留两位小数
    apportioned_areas = {house: round(area * apportionment_factor, 2) for house, area in building_areas.items()}

    return apportioned_areas, apportionment_factor


# 示例使用
building_areas = {
    'House 1': 100,
    'House 2': 150,
    'House 3': 120
}

common_areas = {
    'Lobby': 50,
    'Stairs': 30,
    'Parking': 20
}

apportioned_areas, apportionment_factor = calculate_apportionment(building_areas, common_areas)
print("分摊后的面积:", apportioned_areas)
print("分摊系数:", apportionment_factor)
