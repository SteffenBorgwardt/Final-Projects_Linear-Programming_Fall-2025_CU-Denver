import pandas as pd
from amplpy import AMPL
import matplotlib.pyplot as plt

# run 60 shifts of the single-shift model while updating data in-between iterations
# to simulate a month. accepts list TRUCKS as argument.
def model(TRUCKS):

    # initialize the AMPL environment and choose the solver
    ampl = AMPL()
    ampl.option['solver'] = 'gurobi'

    # Data - note that data is defined here, rather than in an AMPL data file,
    #   so changes to the underlying problem data must be made directly here.

    # Sets
    DEPOTS = ['D1','D2']
    STATIONS = ['S1','S2','S3','S4','S5','S6','S7','S8']
    PRODUCTS = ['P1','P2']

    # Parameters
    # Distance between depots and stations (km)
    distance_raw = {
        'D1': [100, 50, 50, 20, 30, 100, 30, 15],
        'D2': [50, 50, 40, 12, 50, 55, 75, 12]
    }
    distance = {(d, s): distance_raw[d][i] for d in DEPOTS for i, s in enumerate(STATIONS)}
    # Supply available at each depot for each product
    supply_raw = {
        'D1': {'P1': 100000, 'P2': 150000},
        'D2': {'P1': 80000,  'P2': 120000}
    }
    supply = {(d,p): supply_raw[d][p] for d in DEPOTS for p in PRODUCTS}
    # Truck capacity per product
    capacity_truck_raw = {
        'T1':  {'P1': 30, 'P2': 15},
        'T2':  {'P1': 15, 'P2': 30},
        'T3':  {'P1': 30, 'P2': 15},
        'T4':  {'P1': 15, 'P2': 30},
        'T5':  {'P1': 30, 'P2': 15},
        'T6':  {'P1': 15, 'P2': 30},
        'T7':  {'P1': 30, 'P2': 15},
        'T8':  {'P1': 15, 'P2': 30},
        'T9':  {'P1': 30, 'P2': 15},
        'T10': {'P1': 15, 'P2': 30},
        'T11': {'P1': 30, 'P2': 15},
        'T12': {'P1': 15, 'P2': 30},
        'T13': {'P1': 30, 'P2': 15},
        'T14': {'P1': 15, 'P2': 30},
        'T15': {'P1': 30, 'P2': 15},
        'T16': {'P1': 15, 'P2': 30},
        'T17': {'P1': 30, 'P2': 15},
        'T18': {'P1': 15, 'P2': 30},
        'T19': {'P1': 30, 'P2': 15},
        'T20': {'P1': 15, 'P2': 30},
        'T21': {'P1': 30, 'P2': 15},
        'T22': {'P1': 15, 'P2': 30},
        'T23': {'P1': 30, 'P2': 15},
        'T24': {'P1': 15, 'P2': 30}
    }
    capacity_truck = {(t,p): capacity_truck_raw[t][p] for t in TRUCKS for p in PRODUCTS}
    # Currently stored product at each station per product
    capacity_station_raw = {
        'S1': {'P1': 90, 'P2': 110},
        'S2': {'P1': 80, 'P2': 130},
        'S3': {'P1': 90, 'P2': 110},
        'S4': {'P1': 80, 'P2': 130},
        'S5': {'P1': 90, 'P2': 110},
        'S6': {'P1': 80, 'P2': 130},
        'S7': {'P1': 90, 'P2': 110},
        'S8': {'P1': 80, 'P2': 130}
    }
    capacity_station = {(s,p): capacity_station_raw[s][p] for s in STATIONS for p in PRODUCTS}
    # Sales at each station between shifts per product
    sales_station_raw = {
        'S1': {'P1': 45,  'P2': 20},
        'S2': {'P1': 20,  'P2': 13},
        'S3': {'P1': 22,  'P2': 15},
        'S4': {'P1': 25,  'P2': 23},
        'S5': {'P1': 27,  'P2': 30},
        'S6': {'P1': 18,  'P2': 13},
        'S7': {'P1': 16,  'P2': 15},
        'S8': {'P1': 26,  'P2': 15}
    }
    sales_station = {(s,p): sales_station_raw[s][p] for s in STATIONS for p in PRODUCTS}
    # Maximum storage capacity at each station per product
    full_capacity_station_raw = {
        'S1': {'P1': 1800, 'P2': 2200},
        'S2': {'P1': 1600, 'P2': 2600},
        'S3': {'P1': 1800, 'P2': 2200},
        'S4': {'P1': 1600, 'P2': 2600},
        'S5': {'P1': 1800, 'P2': 2200},
        'S6': {'P1': 1600, 'P2': 2600},
        'S7': {'P1': 1800, 'P2': 2200},
        'S8': {'P1': 1600, 'P2': 2600},
    }
    full_capacity_station = {(s,p): full_capacity_station_raw[s][p] for s in STATIONS for p in PRODUCTS}
    # truck driver parameters defaults - to actually change these, you must manually set the value in 
    #   the AMPL python interface. Note that if changing these to specific values, yoou must use a
    #   dictionary like for the above parameters - these defaults just apply to all dictionary keys.
    avg_speed, shift_duration, load_unload_time = 45, 12, 2

    # read auxillary model file
    ampl.read('aux.mod')
    # set auxillary data
    ampl.set['DEPOTS'] = set(DEPOTS)
    ampl.set['STATIONS'] = set(STATIONS)
    ampl.set['PRODUCTS'] = set(PRODUCTS)
    ampl.set['TRUCKS'] = set(TRUCKS)
    ampl.param['distance'] = {(depot, station):num for (depot,station),num in distance.items()}
    ampl.param['capacity_truck'] = {(truck,product):num for (truck,product),num in capacity_truck.items()}
    ampl.param['full_capacity_station'] = {(station,product):num for (station,product),num in full_capacity_station.items()}
    ampl.param['supply'] = {(depot,product):num for (depot,product),num in supply.items()}
    ampl.param['capacity_station'] = {(station,product):num for (station,product),num in capacity_station.items()}
    # solve the auxillary model
    ampl.solve()
    # get the objective value of the auxillary problem
    max_secondary_objective = sum(distance[d,s] * ampl.get_variable('assign_truck')[d,s,t].value() for d in DEPOTS for s in STATIONS for t in TRUCKS)

    # create a list to store solution statistics
    shipped_summary = []

    # reset the AMPL model
    ampl.reset()
    # read the main model file
    ampl.read('proj.mod')
    # set initial data in the model
    ampl.set['DEPOTS'] = set(DEPOTS)
    ampl.set['STATIONS'] = set(STATIONS)
    ampl.set['PRODUCTS'] = set(PRODUCTS)
    ampl.set['TRUCKS'] = set(TRUCKS)
    ampl.param['distance'] = {(depot, station):num for (depot,station),num in distance.items()}
    ampl.param['capacity_truck'] = {(truck,product):num for (truck,product),num in capacity_truck.items()}
    ampl.param['full_capacity_station'] = {(station,product):num for (station,product),num in full_capacity_station.items()}
    ampl.param['max_secondary_objective'] = max_secondary_objective

    # calculate a solution to all 60 shifts in the month
    for iteration in range(60):
        print(f"\n--- Iteration {iteration+1} ---")
        
        # set supply and capacity_station params based of current data
        ampl.param['supply'] = {(depot,product):num for (depot,product),num in supply.items()}
        ampl.param['capacity_station'] = {(station,product):num for (station,product),num in capacity_station.items()}
        # solve the model
        ampl.solve()

        # get the primary objective value
        shipped = {(d, s, p): sum(capacity_truck[t,p] * ampl.get_variable('assign_truck')[d,s,t].value() for t in TRUCKS)
                for d in DEPOTS for s in STATIONS for p in PRODUCTS}

        # Update capacity_station data based on sales and received shipments
        for s in STATIONS:
            for p in PRODUCTS:
                received = sum(shipped[d, s, p] for d in DEPOTS)
                capacity_station[s,p] = max(0, capacity_station[s,p] - sales_station[s,p] + received)

        # Update supply data based on shipments made
        for d in DEPOTS:
            for p in PRODUCTS:
                sent = sum(shipped[d, s, p] for s in STATIONS)
                supply[d,p] = supply[d,p] - sent

        # record solution statistics in shipped_summary
        for p in PRODUCTS:
            total_shipped = sum(shipped[d, s, p] for d in DEPOTS for s in STATIONS)
            total_stock = sum(capacity_station[s,p] for s in STATIONS)
            shipped_summary.append({
                'Iteration': iteration + 1,
                'Product': p,
                'Shipped': total_shipped,
                'Total_Stock': total_stock
            })

    # convert solution statistics list to a pandas DataFrame and return the value
    df_summary = pd.DataFrame(shipped_summary)
    return df_summary

if __name__ == '__main__':

    # define which sets of trucks to be evaluated
    model_dict = {
        # base model
        'M1' : ['T1','T3','T5','T7','T9','T11','T2','T4','T6','T8','T10','T12'],
        # fewer trucks - only type 1 trucks
        'M2' : ['T1','T3','T5','T7','T9','T11'],
        # fewer trucks - only type 2 trucks
        'M3' : ['T2','T4','T6','T8','T10','T12'],
        # same number of trucks - all type 1 trucks
        'M4' : ['T1','T3','T5','T7','T9','T11','T13','T15','T17','T19','T21','T23'],
        # same number of trucks - all type 2 trucks
        'M5' : ['T2','T4','T6','T8','T10','T12','T14','T16','T18','T20','T22','T24'],
        # more trucks - adding type 1 trucks to base
        'M6' : ['T1','T3','T5','T7','T9','T11','T2','T4','T6','T8','T10','T12','T13','T15','T17','T19','T21','T23'],
        #more trucks - adding type 2 trucks to base
        'M7' : ['T1','T3','T5','T7','T9','T11','T2','T4','T6','T8','T10','T12','T14','T16','T18','T20','T22','T24']
    }

    # get list of all different-truck models
    MODELS = model_dict.keys()

    # get solutions to each model
    model_solution_dict = dict()
    for m in MODELS:
        model_solution_dict[m] = model(model_dict[m])

    # Assuming for eaah model DataFrame with columns: Iteration, Product, Total_Stock is returned
    # plot the stock graph of each model for each product
    for product in ['P1','P2']:

        plt.figure(figsize=(10,6))

        for m in MODELS:
            subset = model_solution_dict[m][model_solution_dict[m]['Product'] == product]
            plt.plot(subset['Iteration'], subset['Total_Stock'], marker='o', label=m)

        plt.xlabel('Iteration')
        plt.ylabel('Total Stock')
        plt.title(f'Total Stock of Product {product} Over Iterations')
        plt.legend(title='Model')
        plt.grid(True)
        plt.show()

    # plot the per-iteration shipping graph of model M1
    plt.figure(figsize=(10,6))
    for product in ['P1','P2']:
        subset = model_solution_dict['M1'][model_solution_dict['M1']['Product'] == product]
        plt.plot(subset['Iteration'], subset['Shipped'], marker='o', label=product)
    plt.xlabel('Iteration')
    plt.ylabel('Shipped')
    plt.title(f'Model M1: Shipping Per Iteration')
    plt.legend(title='Product')
    plt.grid(True)
    plt.show()