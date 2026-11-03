import { defineBackend } from '@aws-amplify/backend'
import { Secret } from 'aws-cdk-lib/aws-secretsmanager'
import { Cors, LambdaIntegration, RestApi } from 'aws-cdk-lib/aws-apigateway'
import { booking } from './functions/booking/resource'

const backend = defineBackend({ booking })
const apiStack = backend.createStack('BookingApiStack')
const secret = Secret.fromSecretNameV2(apiStack, 'AzureSecret', 'AzureSecretTerminvereinbarung')
secret.grantRead(backend.booking.resources.lambda)
secret.grantWrite(backend.booking.resources.lambda)

const api = new RestApi(apiStack, 'BookingApi', {
  restApiName: 'terminvereinbarung-api',
  defaultCorsPreflightOptions: {
    allowOrigins: Cors.ALL_ORIGINS,
    allowMethods: Cors.ALL_METHODS,
    allowHeaders: Cors.DEFAULT_HEADERS,
  },
})
const bookingResource = api.root.addResource('booking')
bookingResource.addMethod('GET', new LambdaIntegration(backend.booking.resources.lambda))
bookingResource.addMethod('POST', new LambdaIntegration(backend.booking.resources.lambda))
backend.booking.addEnvironment('OAUTH_REDIRECT_URI', api.urlForPath('/booking'))

backend.addOutput({
  custom: {
    bookingApiUrl: api.urlForPath('/booking'),
  },
})
