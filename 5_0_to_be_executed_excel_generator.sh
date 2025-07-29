#!/bin/bash
#chmod +x 5_0_to_be_executed_excel_generator.sh
#./5_0_to_be_executed_excel_generator.sh

echo "🚀 開始執行 5_1_excel_generator_main.py"
python3 sub_excel_generator_main.py || { echo "❌ generator 執行失敗"; exit 1; }

echo "🚀 開始執行 5_2_excel_path_analyzer.py"
python3 sub_excel_path_analyzer.py || { echo "❌ path analyzer 執行失敗"; exit 1; }

echo "🚀 開始執行 5_3_excel_statistics.py"
python3 sub_excel_statistics.py || { echo "❌ statistics 執行失敗"; exit 1; }

echo "✅ 所有任務完成"
