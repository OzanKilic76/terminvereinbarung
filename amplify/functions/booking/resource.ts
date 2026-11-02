import { defineFunction } from '@aws-amplify/backend'

export const booking = defineFunction({
  name: 'terminvereinbarung-booking',
  entry: './handler.ts',
  timeoutSeconds: 30,
  memoryMB: 512,
  environment: {
    AZURE_SECRET_NAME: 'AzureSecretTerminvereinbarung',
    CALENDAR_OWNER: 'kilic.ozan@outlook.de',
    TIME_ZONE: 'Europe/Berlin',
  },
})
