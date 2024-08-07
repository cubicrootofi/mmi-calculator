import pandas as pd
from sklearn.model_selection import train_test_split
from xgboost import XGBRegressor
from sklearn.metrics import mean_squared_error
import joblib
import matplotlib.pyplot as plt
import numpy as np

# Step 1: Read the Excel file
def read_excel(file_path):
    df = pd.read_excel(file_path)
    return df

# Step 2: Preprocess the data
def preprocess_data(df):
    X = df[['x', 'R', 'q', 'lambda']]
    y = df['y']
    return X, y

# Step 3: Train the XGBoost Regressor with Hyperparameter Tuning
def train_model(X, y):
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.1, random_state=50)
    model = XGBRegressor(n_estimators=2000, learning_rate=0.01, max_depth=4, subsample=0.7, colsample_bytree=0.7)
    model.fit(X_train, y_train, eval_set=[(X_test, y_test)], verbose=False)
    y_pred = model.predict(X_test)
    mse = mean_squared_error(y_test, y_pred)
    print(f"Model trained. Mean Squared Error on test set: {mse}")
    return model

# Step 4: Save the trained model
def save_model(model, file_path):
    joblib.dump(model, file_path)
    print(f"Model saved to {file_path}")

# Step 5: Load the trained model
def load_model(file_path):
    model = joblib.load(file_path)
    return model

# Step 6: Predict using the trained model
def predict(model, x, R, q, lambda_):
    input_data = pd.DataFrame([[x, R, q, lambda_]], columns=['x', 'R', 'q', 'lambda'])
    y_pred = model.predict(input_data)
    return y_pred[0]

# Step 7: Apply Smoothing Function
def smooth_predictions(y_values, alpha=0.1):
    smoothed_values = []
    for i in range(len(y_values)):
        if i == 0:
            smoothed_values.append(y_values[i])
        else:
            smoothed_values.append(alpha * y_values[i] + (1 - alpha) * smoothed_values[i - 1])
    return smoothed_values

# Step 8: Plot the relationship between x and predicted y
def plot_relationship(x_values, y_values):
    plt.figure(figsize=(10, 6))
    plt.plot(x_values, y_values, marker='o', linestyle='-', color='b')
    plt.xlabel('x')
    plt.ylabel('Predicted y')
    plt.title('Relationship between x and Predicted y')
    plt.grid(True)
    plt.show()

# Main script
if __name__ == "__main__":
    # File paths
    excel_file_path = 'sample.xlsx'  # Update this path to your actual Excel file
    model_file_path = 'xgboost_model.pkl'

    # Step 1: Read the data
    df = read_excel(excel_file_path)

    # Step 2: Preprocess the data
    X, y = preprocess_data(df)

    # Step 3: Train the model
    model = train_model(X, y)

    # Step 4: Save the model
    save_model(model, model_file_path)

    # For user input and prediction
    # Load the trained model
    model = load_model(model_file_path)

    # User input for R, q, and lambda
    R_input = 1.5 # float(input("Enter value for R: "))
    q_input = 1.75 # float(input("Enter value for q: "))
    lambda_input = 0.5238 # float(input("Enter value for lambda: "))
    x_value = 6.86
    y_pred = predict(model, x_value, R_input, q_input, lambda_input)
    #smoothed_y_pred = smooth_predictions(y_pred)
    # print(f"{y_pred:.2f}")

    # Predict and store the output for x from 1 to 100
    x_values = np.linspace(1, 100, 100)  # Using np.linspace for smoother plot
    y_values = [predict(model, x, R_input, q_input, lambda_input) for x in x_values]
    #
    # Apply smoothing to the predicted values
    smoothed_y_values = smooth_predictions(y_values)
    #
    # Plot the relationship between x and predicted y
    plot_relationship(x_values, smoothed_y_values)
