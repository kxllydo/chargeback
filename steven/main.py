billing = "../../../../archives/2024/February/billing.csv"
chargeback = "../../../../archives/2024/February/chargeback.xlsx"

from openAI.steven.steps.resourcegroups import step_one

step_one(billing, chargeback)