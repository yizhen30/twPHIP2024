from flask import Flask, request, jsonify
from flask_cors import CORS
import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans

app = Flask(__name__)
CORS(app)

@app.route("/ml_process", methods=["POST"])
def ml_process():
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "請上傳檔案！"}), 400

    df = pd.read_excel(file) if file.filename.endswith(".xlsx") else pd.read_csv(file)

    scaler = StandardScaler()
    scaled_data = scaler.fit_transform(df.iloc[:, 1:])
    
    model = KMeans(n_clusters=3)
    df["cluster"] = model.fit_predict(scaled_data)

    return jsonify(df.to_dict(orient="records"))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
