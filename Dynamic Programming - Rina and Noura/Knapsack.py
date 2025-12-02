i_v_w = [[400, 3], [70, 4], [5, 5]]
max_weight = 10

opt_sol = [[0 for mw in range(max_weight + 1)] for mw in range(len(i_v_w) + 1)]

for sub_item in range(1, len(i_v_w) + 1):
    val, wt = i_v_w[sub_item - 1]
    for sub_weight in range(max_weight + 1):
        if wt > sub_weight:
            opt_sol[sub_item][sub_weight] = opt_sol[sub_item - 1][sub_weight]
        else:
            opt_sol[sub_item][sub_weight] = max(
                opt_sol[sub_item - 1][sub_weight],
                val + opt_sol[sub_item - 1][sub_weight - wt]
            )

for row in opt_sol:
    print(row)

