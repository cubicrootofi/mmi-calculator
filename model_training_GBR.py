import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.ensemble import GradientBoostingRegressor
from sklearn.metrics import mean_squared_error
import joblib

# Step 1: Read the Excel file
def read_excel(file_path):
    df = pd.read_excel(file_path)
    return df

# Step 2: Preprocess the data
def preprocess_data(df):
    X = df[['x', 'R', 'q', 'lambda']]
    y = df['y']
    return X, y

# Step 3: Train the Gradient Boosting Regressor
def train_model(X, y):
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
    model = GradientBoostingRegressor()
    model.fit(X_train, y_train)
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

# Main script
if __name__ == "__main__":
    # File paths
    excel_file_path = 'data.xlsx'  # Update this path to your actual Excel file
    model_file_path = 'gradient_boosting_model.pkl'

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

    # User input
    x_input = float(input("Enter value for x: "))
    R_input = float(input("Enter value for R: "))
    q_input = float(input("Enter value for q: "))
    lambda_input = float(input("Enter value for lambda: "))

    # Predict the output
    y_output = predict(model, x_input, R_input, q_input, lambda_input)
    print(f"Predicted value of y: {y_output}")
