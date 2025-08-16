import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { z } from 'zod';
import commands from '../../commands.js';
import command from './callrecord-get.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';

describe(commands.CALLRECORD_GET, () => {
  const validId = 'e523d2ed-2966-4b6b-925b-754a88034cc5';

  const responseWithSessions = {
    "id": "e523d2ed-2966-4b6b-925b-754a88034cc5",
    "version": 1,
    "type": "groupCall",
    "modalities": [
      "audio"
    ],
    "lastModifiedDateTime": "2025-08-15T12:54:59.8917461Z",
    "startDateTime": "2025-08-15T12:23:32.922748Z",
    "endDateTime": "2025-08-15T12:28:26.0904416Z",
    "joinWebUrl": "https://teams.microsoft.com/l/meetup-join/19%3ameeting_MGU5Mjk3MjgtMjMyZS00ZGEzLTk2YjktZGNlMTc4NjljMmYz%40thread.v2/0?context=%7b%22Tid%22%3a%229d66187e-13f0-4666-9bac-be67ddd4b676%22%2c%22Oid%22%3a%2242559007-03c6-42c8-971f-cb79fd381a5a%22%7d",
    "organizer": {
      "acsUser": null,
      "spoolUser": null,
      "phone": null,
      "guest": null,
      "encrypted": null,
      "onPremises": null,
      "acsApplicationInstance": null,
      "spoolApplicationInstance": null,
      "applicationInstance": null,
      "application": null,
      "device": null,
      "user": {
        "id": "42559007-03c6-42c8-971f-cb79fd381a5a",
        "displayName": "John Doe",
        "tenantId": "9d66187e-13f0-4666-9bac-be67ddd4b676"
      }
    },
    "participants": [
      {
        "acsUser": null,
        "spoolUser": null,
        "phone": null,
        "guest": null,
        "encrypted": null,
        "onPremises": null,
        "acsApplicationInstance": null,
        "spoolApplicationInstance": null,
        "applicationInstance": null,
        "application": null,
        "device": null,
        "user": {
          "id": "42559007-03c6-42c8-971f-cb79fd381a5a",
          "displayName": "John Doe",
          "tenantId": "9d66187e-13f0-4666-9bac-be67ddd4b676"
        }
      }
    ],
    "organizer_v2": {
      "id": "42559007-03c6-42c8-971f-cb79fd381a5a",
      "identity": {
        "endpointType": null,
        "acsUser": null,
        "spoolUser": null,
        "phone": null,
        "guest": null,
        "encrypted": null,
        "onPremises": null,
        "acsApplicationInstance": null,
        "spoolApplicationInstance": null,
        "applicationInstance": null,
        "application": null,
        "device": null,
        "azureCommunicationServicesUser": null,
        "assertedIdentity": null,
        "user": {
          "id": "42559007-03c6-42c8-971f-cb79fd381a5a",
          "displayName": "John Doe",
          "tenantId": "9d66187e-13f0-4666-9bac-be67ddd4b676",
          "userPrincipalName": "john.doe@contoso.com"
        }
      },
      "administrativeUnitInfos": [
        {
          "id": "6d3eb178-988b-430b-974f-1254113e4522"
        }
      ]
    },
    "sessions": [
      {
        "id": "4181f0b0-bbc5-4ff3-ad07-8a496bcada10",
        "modalities": [
          "audio"
        ],
        "startDateTime": "2025-08-15T12:23:32.922748Z",
        "endDateTime": "2025-08-15T12:28:26.0904416Z",
        "isTest": false,
        "failureInfo": null,
        "caller": {
          "name": "JOHNDOE-PC",
          "cpuName": "12th Gen Intel(R) Core(TM) i9-12900H",
          "cpuCoresCount": 14,
          "cpuProcessorSpeedInMhz": 1800,
          "userAgent": {
            "headerValue": "releases/CL2025.R25",
            "applicationVersion": null,
            "platform": "windows",
            "productFamily": "teams",
            "communicationServiceId": null,
            "azureADAppId": null
          },
          "identity": {
            "acsUser": null,
            "spoolUser": null,
            "phone": null,
            "guest": null,
            "encrypted": null,
            "onPremises": null,
            "acsApplicationInstance": null,
            "spoolApplicationInstance": null,
            "applicationInstance": null,
            "application": null,
            "device": null,
            "user": {
              "id": "42559007-03c6-42c8-971f-cb79fd381a5a",
              "displayName": "John Doe",
              "tenantId": "9d66187e-13f0-4666-9bac-be67ddd4b676"
            }
          },
          "associatedIdentity": {
            "id": "42559007-03c6-42c8-971f-cb79fd381a5a",
            "displayName": "John Doe",
            "tenantId": "9d66187e-13f0-4666-9bac-be67ddd4b676",
            "userPrincipalName": "john.doe@contoso.com"
          },
          "feedback": {
            "text": null,
            "rating": "notRated",
            "tokens": {}
          }
        },
        "callee": {
          "userAgent": {
            "headerValue": null,
            "applicationVersion": null,
            "platform": "unknown",
            "productFamily": "unknown",
            "communicationServiceId": null,
            "azureADAppId": null
          }
        },
        "segments": [
          {
            "id": "4181f0b0-bbc5-4ff3-ad07-8a496bcada10",
            "startDateTime": "2025-08-15T12:23:32.922748Z",
            "endDateTime": "2025-08-15T12:28:26.0904416Z",
            "failureInfo": null,
            "caller": {
              "name": "JOHNDOE-PC",
              "cpuName": "12th Gen Intel(R) Core(TM) i9-12900H",
              "cpuCoresCount": 14,
              "cpuProcessorSpeedInMhz": 1800,
              "userAgent": {
                "headerValue": "releases/CL2025.R25",
                "applicationVersion": null,
                "platform": "windows",
                "productFamily": "teams",
                "communicationServiceId": null,
                "azureADAppId": null
              },
              "identity": {
                "acsUser": null,
                "spoolUser": null,
                "phone": null,
                "guest": null,
                "encrypted": null,
                "onPremises": null,
                "acsApplicationInstance": null,
                "spoolApplicationInstance": null,
                "applicationInstance": null,
                "application": null,
                "device": null,
                "user": {
                  "id": "42559007-03c6-42c8-971f-cb79fd381a5a",
                  "displayName": "John Doe",
                  "tenantId": "9d66187e-13f0-4666-9bac-be67ddd4b676"
                }
              },
              "associatedIdentity": {
                "id": "42559007-03c6-42c8-971f-cb79fd381a5a",
                "displayName": "John Doe",
                "tenantId": "9d66187e-13f0-4666-9bac-be67ddd4b676",
                "userPrincipalName": "john.doe@contoso.com"
              },
              "feedback": {
                "text": null,
                "rating": "notRated",
                "tokens": {}
              }
            },
            "callee": {
              "userAgent": {
                "headerValue": null,
                "applicationVersion": null,
                "platform": "unknown",
                "productFamily": "unknown",
                "communicationServiceId": null,
                "azureADAppId": null
              }
            },
            "media": [
              {
                "label": "data",
                "callerNetwork": {
                  "ipAddress": "192.168.0.243",
                  "subnet": "192.168.0.0",
                  "linkSpeed": 526500000,
                  "connectionType": "wifi",
                  "port": 50048,
                  "reflexiveIPAddress": "",
                  "relayIPAddress": "",
                  "relayPort": 3481,
                  "macAddress": "",
                  "wifiMicrosoftDriver": "",
                  "wifiMicrosoftDriverVersion": "Microsoft:10.0.26100.4484",
                  "wifiVendorDriver": "Intel(R) Wi-Fi 6E AX211 160MHz",
                  "wifiVendorDriverVersion": "Intel:23.110.0.5",
                  "wifiChannel": 36,
                  "wifiBand": "frequency50GHz",
                  "basicServiceSetIdentifier": "",
                  "wifiRadioType": "wifi80211ac",
                  "wifiSignalStrength": 94,
                  "wifiBatteryCharge": 100,
                  "dnsSuffix": "",
                  "sentQualityEventRatio": null,
                  "receivedQualityEventRatio": null,
                  "delayEventRatio": null,
                  "bandwidthLowEventRatio": null,
                  "networkTransportProtocol": "udp",
                  "traceRouteHops": []
                },
                "calleeNetwork": {
                  "ipAddress": "",
                  "subnet": null,
                  "linkSpeed": 0,
                  "connectionType": "wired",
                  "port": 3481,
                  "reflexiveIPAddress": "",
                  "relayIPAddress": null,
                  "relayPort": null,
                  "macAddress": null,
                  "wifiMicrosoftDriver": null,
                  "wifiMicrosoftDriverVersion": null,
                  "wifiVendorDriver": null,
                  "wifiVendorDriverVersion": null,
                  "wifiChannel": null,
                  "wifiBand": "unknown",
                  "basicServiceSetIdentifier": null,
                  "wifiRadioType": "unknown",
                  "wifiSignalStrength": null,
                  "wifiBatteryCharge": null,
                  "dnsSuffix": null,
                  "sentQualityEventRatio": null,
                  "receivedQualityEventRatio": null,
                  "delayEventRatio": null,
                  "bandwidthLowEventRatio": null,
                  "networkTransportProtocol": "udp",
                  "traceRouteHops": []
                },
                "callerDevice": {
                  "captureDeviceName": null,
                  "captureDeviceDriver": null,
                  "renderDeviceName": null,
                  "renderDeviceDriver": null,
                  "sentSignalLevel": null,
                  "receivedSignalLevel": null,
                  "sentNoiseLevel": null,
                  "receivedNoiseLevel": null,
                  "initialSignalLevelRootMeanSquare": null,
                  "cpuInsufficentEventRatio": null,
                  "renderNotFunctioningEventRatio": null,
                  "captureNotFunctioningEventRatio": null,
                  "deviceGlitchEventRatio": null,
                  "lowSpeechToNoiseEventRatio": null,
                  "lowSpeechLevelEventRatio": null,
                  "deviceClippingEventRatio": null,
                  "howlingEventCount": null,
                  "renderZeroVolumeEventRatio": null,
                  "renderMuteEventRatio": null,
                  "micGlitchRate": null,
                  "speakerGlitchRate": null
                },
                "calleeDevice": {
                  "captureDeviceName": null,
                  "captureDeviceDriver": null,
                  "renderDeviceName": null,
                  "renderDeviceDriver": null,
                  "sentSignalLevel": null,
                  "receivedSignalLevel": null,
                  "sentNoiseLevel": null,
                  "receivedNoiseLevel": null,
                  "initialSignalLevelRootMeanSquare": null,
                  "cpuInsufficentEventRatio": null,
                  "renderNotFunctioningEventRatio": null,
                  "captureNotFunctioningEventRatio": null,
                  "deviceGlitchEventRatio": null,
                  "lowSpeechToNoiseEventRatio": null,
                  "lowSpeechLevelEventRatio": null,
                  "deviceClippingEventRatio": null,
                  "howlingEventCount": null,
                  "renderZeroVolumeEventRatio": null,
                  "renderMuteEventRatio": null,
                  "micGlitchRate": null,
                  "speakerGlitchRate": null
                },
                "streams": [
                  {
                    "streamId": "21549",
                    "startDateTime": null,
                    "endDateTime": null,
                    "streamDirection": "callerToCallee",
                    "averageAudioDegradation": null,
                    "averageJitter": null,
                    "maxJitter": null,
                    "averagePacketLossRate": null,
                    "maxPacketLossRate": null,
                    "averageRatioOfConcealedSamples": null,
                    "maxRatioOfConcealedSamples": null,
                    "averageRoundTripTime": null,
                    "maxRoundTripTime": null,
                    "packetUtilization": 0,
                    "averageBandwidthEstimate": null,
                    "wasMediaBypassed": null,
                    "postForwardErrorCorrectionPacketLossRate": null,
                    "averageVideoFrameLossPercentage": null,
                    "averageReceivedFrameRate": null,
                    "lowFrameRateRatio": null,
                    "averageVideoPacketLossRate": null,
                    "averageVideoFrameRate": null,
                    "lowVideoProcessingCapabilityRatio": null,
                    "averageAudioNetworkJitter": null,
                    "maxAudioNetworkJitter": null,
                    "audioCodec": "unknown",
                    "videoCodec": "unknown",
                    "rmsFreezeDuration": null,
                    "averageFreezeDuration": null,
                    "isAudioForwardErrorCorrectionUsed": null
                  },
                  {
                    "streamId": "22901",
                    "startDateTime": null,
                    "endDateTime": null,
                    "streamDirection": "calleeToCaller",
                    "averageAudioDegradation": null,
                    "averageJitter": null,
                    "maxJitter": null,
                    "averagePacketLossRate": null,
                    "maxPacketLossRate": null,
                    "averageRatioOfConcealedSamples": null,
                    "maxRatioOfConcealedSamples": null,
                    "averageRoundTripTime": null,
                    "maxRoundTripTime": null,
                    "packetUtilization": 0,
                    "averageBandwidthEstimate": null,
                    "wasMediaBypassed": null,
                    "postForwardErrorCorrectionPacketLossRate": null,
                    "averageVideoFrameLossPercentage": null,
                    "averageReceivedFrameRate": null,
                    "lowFrameRateRatio": null,
                    "averageVideoPacketLossRate": null,
                    "averageVideoFrameRate": null,
                    "lowVideoProcessingCapabilityRatio": null,
                    "averageAudioNetworkJitter": null,
                    "maxAudioNetworkJitter": null,
                    "audioCodec": "unknown",
                    "videoCodec": "unknown",
                    "rmsFreezeDuration": null,
                    "averageFreezeDuration": null,
                    "isAudioForwardErrorCorrectionUsed": null
                  }
                ]
              },
              {
                "label": "main-audio",
                "callerNetwork": {
                  "ipAddress": "192.168.0.243",
                  "subnet": "192.168.0.0",
                  "linkSpeed": 526500000,
                  "connectionType": "wifi",
                  "port": 50005,
                  "reflexiveIPAddress": "",
                  "relayIPAddress": "",
                  "relayPort": 3479,
                  "macAddress": "",
                  "wifiMicrosoftDriver": "",
                  "wifiMicrosoftDriverVersion": "Microsoft:10.0.26100.4484",
                  "wifiVendorDriver": "Intel(R) Wi-Fi 6E AX211 160MHz",
                  "wifiVendorDriverVersion": "Intel:23.110.0.5",
                  "wifiChannel": 36,
                  "wifiBand": "frequency50GHz",
                  "basicServiceSetIdentifier": "",
                  "wifiRadioType": "wifi80211ac",
                  "wifiSignalStrength": 94,
                  "wifiBatteryCharge": 100,
                  "dnsSuffix": "",
                  "sentQualityEventRatio": 0,
                  "receivedQualityEventRatio": 0,
                  "delayEventRatio": null,
                  "bandwidthLowEventRatio": null,
                  "networkTransportProtocol": "udp",
                  "traceRouteHops": []
                },
                "calleeNetwork": {
                  "ipAddress": "",
                  "subnet": null,
                  "linkSpeed": 0,
                  "connectionType": "wired",
                  "port": 3479,
                  "reflexiveIPAddress": "",
                  "relayIPAddress": null,
                  "relayPort": null,
                  "macAddress": null,
                  "wifiMicrosoftDriver": null,
                  "wifiMicrosoftDriverVersion": null,
                  "wifiVendorDriver": null,
                  "wifiVendorDriverVersion": null,
                  "wifiChannel": null,
                  "wifiBand": "unknown",
                  "basicServiceSetIdentifier": null,
                  "wifiRadioType": "unknown",
                  "wifiSignalStrength": null,
                  "wifiBatteryCharge": null,
                  "dnsSuffix": null,
                  "sentQualityEventRatio": null,
                  "receivedQualityEventRatio": null,
                  "delayEventRatio": null,
                  "bandwidthLowEventRatio": null,
                  "networkTransportProtocol": "udp",
                  "traceRouteHops": []
                },
                "callerDevice": {
                  "captureDeviceName": "Realtek(R) Audio",
                  "captureDeviceDriver": "Realtek Semiconductor Corp.: 6.0.9780.1",
                  "renderDeviceName": "Realtek(R) Audio",
                  "renderDeviceDriver": "Realtek Semiconductor Corp.: 6.0.9780.1",
                  "sentSignalLevel": -23,
                  "receivedSignalLevel": null,
                  "sentNoiseLevel": -70,
                  "receivedNoiseLevel": null,
                  "initialSignalLevelRootMeanSquare": null,
                  "cpuInsufficentEventRatio": 0,
                  "renderNotFunctioningEventRatio": 0,
                  "captureNotFunctioningEventRatio": 0,
                  "deviceGlitchEventRatio": 0,
                  "lowSpeechToNoiseEventRatio": 0,
                  "lowSpeechLevelEventRatio": 0,
                  "deviceClippingEventRatio": 0,
                  "howlingEventCount": 0,
                  "renderZeroVolumeEventRatio": 0,
                  "renderMuteEventRatio": 0,
                  "micGlitchRate": null,
                  "speakerGlitchRate": null
                },
                "calleeDevice": {
                  "captureDeviceName": null,
                  "captureDeviceDriver": null,
                  "renderDeviceName": null,
                  "renderDeviceDriver": null,
                  "sentSignalLevel": null,
                  "receivedSignalLevel": null,
                  "sentNoiseLevel": null,
                  "receivedNoiseLevel": null,
                  "initialSignalLevelRootMeanSquare": null,
                  "cpuInsufficentEventRatio": null,
                  "renderNotFunctioningEventRatio": null,
                  "captureNotFunctioningEventRatio": null,
                  "deviceGlitchEventRatio": null,
                  "lowSpeechToNoiseEventRatio": null,
                  "lowSpeechLevelEventRatio": null,
                  "deviceClippingEventRatio": null,
                  "howlingEventCount": 0,
                  "renderZeroVolumeEventRatio": null,
                  "renderMuteEventRatio": null,
                  "micGlitchRate": null,
                  "speakerGlitchRate": null
                },
                "streams": [
                  {
                    "streamId": "648",
                    "startDateTime": null,
                    "endDateTime": null,
                    "streamDirection": "callerToCallee",
                    "averageAudioDegradation": 0,
                    "averageJitter": "PT0.002S",
                    "maxJitter": "PT0.01S",
                    "averagePacketLossRate": 0,
                    "maxPacketLossRate": 0,
                    "averageRatioOfConcealedSamples": 0.008604,
                    "maxRatioOfConcealedSamples": null,
                    "averageRoundTripTime": "PT0.047S",
                    "maxRoundTripTime": "PT0.049S",
                    "packetUtilization": 529,
                    "averageBandwidthEstimate": 2477474,
                    "wasMediaBypassed": null,
                    "postForwardErrorCorrectionPacketLossRate": null,
                    "averageVideoFrameLossPercentage": null,
                    "averageReceivedFrameRate": null,
                    "lowFrameRateRatio": null,
                    "averageVideoPacketLossRate": null,
                    "averageVideoFrameRate": null,
                    "lowVideoProcessingCapabilityRatio": null,
                    "averageAudioNetworkJitter": "PT0.008S",
                    "maxAudioNetworkJitter": "PT0.029S",
                    "audioCodec": "satin",
                    "videoCodec": "unknown",
                    "rmsFreezeDuration": null,
                    "averageFreezeDuration": null,
                    "isAudioForwardErrorCorrectionUsed": false
                  },
                  {
                    "streamId": "1000",
                    "startDateTime": null,
                    "endDateTime": null,
                    "streamDirection": "calleeToCaller",
                    "averageAudioDegradation": null,
                    "averageJitter": null,
                    "maxJitter": null,
                    "averagePacketLossRate": null,
                    "maxPacketLossRate": null,
                    "averageRatioOfConcealedSamples": null,
                    "maxRatioOfConcealedSamples": null,
                    "averageRoundTripTime": null,
                    "maxRoundTripTime": null,
                    "packetUtilization": 0,
                    "averageBandwidthEstimate": null,
                    "wasMediaBypassed": null,
                    "postForwardErrorCorrectionPacketLossRate": null,
                    "averageVideoFrameLossPercentage": null,
                    "averageReceivedFrameRate": null,
                    "lowFrameRateRatio": null,
                    "averageVideoPacketLossRate": null,
                    "averageVideoFrameRate": null,
                    "lowVideoProcessingCapabilityRatio": null,
                    "averageAudioNetworkJitter": null,
                    "maxAudioNetworkJitter": null,
                    "audioCodec": "unknown",
                    "videoCodec": "unknown",
                    "rmsFreezeDuration": null,
                    "averageFreezeDuration": null,
                    "isAudioForwardErrorCorrectionUsed": null
                  }
                ]
              }
            ]
          }
        ]
      }
    ]
  };

  const responseWithParticipants = {
    "id": "e523d2ed-2966-4b6b-925b-754a88034cc5",
    "participants_v2": [
      {
        "id": "42559007-03c6-42c8-971f-cb79fd381a5a",
        "identity": {
          "endpointType": null,
          "acsUser": null,
          "spoolUser": null,
          "phone": null,
          "guest": null,
          "encrypted": null,
          "onPremises": null,
          "acsApplicationInstance": null,
          "spoolApplicationInstance": null,
          "applicationInstance": null,
          "application": null,
          "device": null,
          "azureCommunicationServicesUser": null,
          "assertedIdentity": null,
          "user": {
            "id": "42559007-03c6-42c8-971f-cb79fd381a5a",
            "displayName": "John Doe",
            "tenantId": "9d66187e-13f0-4666-9bac-be67ddd4b676",
            "userPrincipalName": "john.doe@contoso.com"
          }
        },
        "administrativeUnitInfos": [
          {
            "id": "6d3eb178-988b-430b-974f-1254113e4522"
          }
        ]
      }
    ]
  };

  const response = {
    "id": "e523d2ed-2966-4b6b-925b-754a88034cc5",
    "version": 1,
    "type": "groupCall",
    "modalities": [
      "audio"
    ],
    "lastModifiedDateTime": "2025-08-15T12:54:59.8917461Z",
    "startDateTime": "2025-08-15T12:23:32.922748Z",
    "endDateTime": "2025-08-15T12:28:26.0904416Z",
    "joinWebUrl": "https://teams.microsoft.com/l/meetup-join/19%3ameeting_MGU5Mjk3MjgtMjMyZS00ZGEzLTk2YjktZGNlMTc4NjljMmYz%40thread.v2/0?context=%7b%22Tid%22%3a%229d66187e-13f0-4666-9bac-be67ddd4b676%22%2c%22Oid%22%3a%2242559007-03c6-42c8-971f-cb79fd381a5a%22%7d",
    "organizer": {
      "acsUser": null,
      "spoolUser": null,
      "phone": null,
      "guest": null,
      "encrypted": null,
      "onPremises": null,
      "acsApplicationInstance": null,
      "spoolApplicationInstance": null,
      "applicationInstance": null,
      "application": null,
      "device": null,
      "user": {
        "id": "42559007-03c6-42c8-971f-cb79fd381a5a",
        "displayName": "John Doe",
        "tenantId": "9d66187e-13f0-4666-9bac-be67ddd4b676"
      }
    },
    "participants": [
      {
        "acsUser": null,
        "spoolUser": null,
        "phone": null,
        "guest": null,
        "encrypted": null,
        "onPremises": null,
        "acsApplicationInstance": null,
        "spoolApplicationInstance": null,
        "applicationInstance": null,
        "application": null,
        "device": null,
        "user": {
          "id": "42559007-03c6-42c8-971f-cb79fd381a5a",
          "displayName": "John Doe",
          "tenantId": "9d66187e-13f0-4666-9bac-be67ddd4b676"
        }
      }
    ],
    "organizer_v2": {
      "id": "42559007-03c6-42c8-971f-cb79fd381a5a",
      "identity": {
        "endpointType": null,
        "acsUser": null,
        "spoolUser": null,
        "phone": null,
        "guest": null,
        "encrypted": null,
        "onPremises": null,
        "acsApplicationInstance": null,
        "spoolApplicationInstance": null,
        "applicationInstance": null,
        "application": null,
        "device": null,
        "azureCommunicationServicesUser": null,
        "assertedIdentity": null,
        "user": {
          "id": "42559007-03c6-42c8-971f-cb79fd381a5a",
          "displayName": "John Doe",
          "tenantId": "9d66187e-13f0-4666-9bac-be67ddd4b676",
          "userPrincipalName": "john.doe@contoso.com"
        }
      },
      "administrativeUnitInfos": [
        {
          "id": "6d3eb178-988b-430b-974f-1254113e4522"
        }
      ]
    },
    "sessions": [
      {
        "id": "4181f0b0-bbc5-4ff3-ad07-8a496bcada10",
        "modalities": [
          "audio"
        ],
        "startDateTime": "2025-08-15T12:23:32.922748Z",
        "endDateTime": "2025-08-15T12:28:26.0904416Z",
        "isTest": false,
        "failureInfo": null,
        "caller": {
          "name": "JOHNDOE-PC",
          "cpuName": "12th Gen Intel(R) Core(TM) i9-12900H",
          "cpuCoresCount": 14,
          "cpuProcessorSpeedInMhz": 1800,
          "userAgent": {
            "headerValue": "releases/CL2025.R25",
            "applicationVersion": null,
            "platform": "windows",
            "productFamily": "teams",
            "communicationServiceId": null,
            "azureADAppId": null
          },
          "identity": {
            "acsUser": null,
            "spoolUser": null,
            "phone": null,
            "guest": null,
            "encrypted": null,
            "onPremises": null,
            "acsApplicationInstance": null,
            "spoolApplicationInstance": null,
            "applicationInstance": null,
            "application": null,
            "device": null,
            "user": {
              "id": "42559007-03c6-42c8-971f-cb79fd381a5a",
              "displayName": "John Doe",
              "tenantId": "9d66187e-13f0-4666-9bac-be67ddd4b676"
            }
          },
          "associatedIdentity": {
            "id": "42559007-03c6-42c8-971f-cb79fd381a5a",
            "displayName": "John Doe",
            "tenantId": "9d66187e-13f0-4666-9bac-be67ddd4b676",
            "userPrincipalName": "john.doe@contoso.com"
          },
          "feedback": {
            "text": null,
            "rating": "notRated",
            "tokens": {}
          }
        },
        "callee": {
          "userAgent": {
            "headerValue": null,
            "applicationVersion": null,
            "platform": "unknown",
            "productFamily": "unknown",
            "communicationServiceId": null,
            "azureADAppId": null
          }
        },
        "segments": [
          {
            "id": "4181f0b0-bbc5-4ff3-ad07-8a496bcada10",
            "startDateTime": "2025-08-15T12:23:32.922748Z",
            "endDateTime": "2025-08-15T12:28:26.0904416Z",
            "failureInfo": null,
            "caller": {
              "name": "JOHNDOE-PC",
              "cpuName": "12th Gen Intel(R) Core(TM) i9-12900H",
              "cpuCoresCount": 14,
              "cpuProcessorSpeedInMhz": 1800,
              "userAgent": {
                "headerValue": "releases/CL2025.R25",
                "applicationVersion": null,
                "platform": "windows",
                "productFamily": "teams",
                "communicationServiceId": null,
                "azureADAppId": null
              },
              "identity": {
                "acsUser": null,
                "spoolUser": null,
                "phone": null,
                "guest": null,
                "encrypted": null,
                "onPremises": null,
                "acsApplicationInstance": null,
                "spoolApplicationInstance": null,
                "applicationInstance": null,
                "application": null,
                "device": null,
                "user": {
                  "id": "42559007-03c6-42c8-971f-cb79fd381a5a",
                  "displayName": "John Doe",
                  "tenantId": "9d66187e-13f0-4666-9bac-be67ddd4b676"
                }
              },
              "associatedIdentity": {
                "id": "42559007-03c6-42c8-971f-cb79fd381a5a",
                "displayName": "John Doe",
                "tenantId": "9d66187e-13f0-4666-9bac-be67ddd4b676",
                "userPrincipalName": "john.doe@contoso.com"
              },
              "feedback": {
                "text": null,
                "rating": "notRated",
                "tokens": {}
              }
            },
            "callee": {
              "userAgent": {
                "headerValue": null,
                "applicationVersion": null,
                "platform": "unknown",
                "productFamily": "unknown",
                "communicationServiceId": null,
                "azureADAppId": null
              }
            },
            "media": [
              {
                "label": "data",
                "callerNetwork": {
                  "ipAddress": "192.168.0.243",
                  "subnet": "192.168.0.0",
                  "linkSpeed": 526500000,
                  "connectionType": "wifi",
                  "port": 50048,
                  "reflexiveIPAddress": "",
                  "relayIPAddress": "",
                  "relayPort": 3481,
                  "macAddress": "",
                  "wifiMicrosoftDriver": "",
                  "wifiMicrosoftDriverVersion": "Microsoft:10.0.26100.4484",
                  "wifiVendorDriver": "Intel(R) Wi-Fi 6E AX211 160MHz",
                  "wifiVendorDriverVersion": "Intel:23.110.0.5",
                  "wifiChannel": 36,
                  "wifiBand": "frequency50GHz",
                  "basicServiceSetIdentifier": "",
                  "wifiRadioType": "wifi80211ac",
                  "wifiSignalStrength": 94,
                  "wifiBatteryCharge": 100,
                  "dnsSuffix": "",
                  "sentQualityEventRatio": null,
                  "receivedQualityEventRatio": null,
                  "delayEventRatio": null,
                  "bandwidthLowEventRatio": null,
                  "networkTransportProtocol": "udp",
                  "traceRouteHops": []
                },
                "calleeNetwork": {
                  "ipAddress": "",
                  "subnet": null,
                  "linkSpeed": 0,
                  "connectionType": "wired",
                  "port": 3481,
                  "reflexiveIPAddress": "",
                  "relayIPAddress": null,
                  "relayPort": null,
                  "macAddress": null,
                  "wifiMicrosoftDriver": null,
                  "wifiMicrosoftDriverVersion": null,
                  "wifiVendorDriver": null,
                  "wifiVendorDriverVersion": null,
                  "wifiChannel": null,
                  "wifiBand": "unknown",
                  "basicServiceSetIdentifier": null,
                  "wifiRadioType": "unknown",
                  "wifiSignalStrength": null,
                  "wifiBatteryCharge": null,
                  "dnsSuffix": null,
                  "sentQualityEventRatio": null,
                  "receivedQualityEventRatio": null,
                  "delayEventRatio": null,
                  "bandwidthLowEventRatio": null,
                  "networkTransportProtocol": "udp",
                  "traceRouteHops": []
                },
                "callerDevice": {
                  "captureDeviceName": null,
                  "captureDeviceDriver": null,
                  "renderDeviceName": null,
                  "renderDeviceDriver": null,
                  "sentSignalLevel": null,
                  "receivedSignalLevel": null,
                  "sentNoiseLevel": null,
                  "receivedNoiseLevel": null,
                  "initialSignalLevelRootMeanSquare": null,
                  "cpuInsufficentEventRatio": null,
                  "renderNotFunctioningEventRatio": null,
                  "captureNotFunctioningEventRatio": null,
                  "deviceGlitchEventRatio": null,
                  "lowSpeechToNoiseEventRatio": null,
                  "lowSpeechLevelEventRatio": null,
                  "deviceClippingEventRatio": null,
                  "howlingEventCount": null,
                  "renderZeroVolumeEventRatio": null,
                  "renderMuteEventRatio": null,
                  "micGlitchRate": null,
                  "speakerGlitchRate": null
                },
                "calleeDevice": {
                  "captureDeviceName": null,
                  "captureDeviceDriver": null,
                  "renderDeviceName": null,
                  "renderDeviceDriver": null,
                  "sentSignalLevel": null,
                  "receivedSignalLevel": null,
                  "sentNoiseLevel": null,
                  "receivedNoiseLevel": null,
                  "initialSignalLevelRootMeanSquare": null,
                  "cpuInsufficentEventRatio": null,
                  "renderNotFunctioningEventRatio": null,
                  "captureNotFunctioningEventRatio": null,
                  "deviceGlitchEventRatio": null,
                  "lowSpeechToNoiseEventRatio": null,
                  "lowSpeechLevelEventRatio": null,
                  "deviceClippingEventRatio": null,
                  "howlingEventCount": null,
                  "renderZeroVolumeEventRatio": null,
                  "renderMuteEventRatio": null,
                  "micGlitchRate": null,
                  "speakerGlitchRate": null
                },
                "streams": [
                  {
                    "streamId": "21549",
                    "startDateTime": null,
                    "endDateTime": null,
                    "streamDirection": "callerToCallee",
                    "averageAudioDegradation": null,
                    "averageJitter": null,
                    "maxJitter": null,
                    "averagePacketLossRate": null,
                    "maxPacketLossRate": null,
                    "averageRatioOfConcealedSamples": null,
                    "maxRatioOfConcealedSamples": null,
                    "averageRoundTripTime": null,
                    "maxRoundTripTime": null,
                    "packetUtilization": 0,
                    "averageBandwidthEstimate": null,
                    "wasMediaBypassed": null,
                    "postForwardErrorCorrectionPacketLossRate": null,
                    "averageVideoFrameLossPercentage": null,
                    "averageReceivedFrameRate": null,
                    "lowFrameRateRatio": null,
                    "averageVideoPacketLossRate": null,
                    "averageVideoFrameRate": null,
                    "lowVideoProcessingCapabilityRatio": null,
                    "averageAudioNetworkJitter": null,
                    "maxAudioNetworkJitter": null,
                    "audioCodec": "unknown",
                    "videoCodec": "unknown",
                    "rmsFreezeDuration": null,
                    "averageFreezeDuration": null,
                    "isAudioForwardErrorCorrectionUsed": null
                  },
                  {
                    "streamId": "22901",
                    "startDateTime": null,
                    "endDateTime": null,
                    "streamDirection": "calleeToCaller",
                    "averageAudioDegradation": null,
                    "averageJitter": null,
                    "maxJitter": null,
                    "averagePacketLossRate": null,
                    "maxPacketLossRate": null,
                    "averageRatioOfConcealedSamples": null,
                    "maxRatioOfConcealedSamples": null,
                    "averageRoundTripTime": null,
                    "maxRoundTripTime": null,
                    "packetUtilization": 0,
                    "averageBandwidthEstimate": null,
                    "wasMediaBypassed": null,
                    "postForwardErrorCorrectionPacketLossRate": null,
                    "averageVideoFrameLossPercentage": null,
                    "averageReceivedFrameRate": null,
                    "lowFrameRateRatio": null,
                    "averageVideoPacketLossRate": null,
                    "averageVideoFrameRate": null,
                    "lowVideoProcessingCapabilityRatio": null,
                    "averageAudioNetworkJitter": null,
                    "maxAudioNetworkJitter": null,
                    "audioCodec": "unknown",
                    "videoCodec": "unknown",
                    "rmsFreezeDuration": null,
                    "averageFreezeDuration": null,
                    "isAudioForwardErrorCorrectionUsed": null
                  }
                ]
              },
              {
                "label": "main-audio",
                "callerNetwork": {
                  "ipAddress": "192.168.0.243",
                  "subnet": "192.168.0.0",
                  "linkSpeed": 526500000,
                  "connectionType": "wifi",
                  "port": 50005,
                  "reflexiveIPAddress": "",
                  "relayIPAddress": "",
                  "relayPort": 3479,
                  "macAddress": "",
                  "wifiMicrosoftDriver": "",
                  "wifiMicrosoftDriverVersion": "Microsoft:10.0.26100.4484",
                  "wifiVendorDriver": "Intel(R) Wi-Fi 6E AX211 160MHz",
                  "wifiVendorDriverVersion": "Intel:23.110.0.5",
                  "wifiChannel": 36,
                  "wifiBand": "frequency50GHz",
                  "basicServiceSetIdentifier": "",
                  "wifiRadioType": "wifi80211ac",
                  "wifiSignalStrength": 94,
                  "wifiBatteryCharge": 100,
                  "dnsSuffix": "",
                  "sentQualityEventRatio": 0,
                  "receivedQualityEventRatio": 0,
                  "delayEventRatio": null,
                  "bandwidthLowEventRatio": null,
                  "networkTransportProtocol": "udp",
                  "traceRouteHops": []
                },
                "calleeNetwork": {
                  "ipAddress": "",
                  "subnet": null,
                  "linkSpeed": 0,
                  "connectionType": "wired",
                  "port": 3479,
                  "reflexiveIPAddress": "",
                  "relayIPAddress": null,
                  "relayPort": null,
                  "macAddress": null,
                  "wifiMicrosoftDriver": null,
                  "wifiMicrosoftDriverVersion": null,
                  "wifiVendorDriver": null,
                  "wifiVendorDriverVersion": null,
                  "wifiChannel": null,
                  "wifiBand": "unknown",
                  "basicServiceSetIdentifier": null,
                  "wifiRadioType": "unknown",
                  "wifiSignalStrength": null,
                  "wifiBatteryCharge": null,
                  "dnsSuffix": null,
                  "sentQualityEventRatio": null,
                  "receivedQualityEventRatio": null,
                  "delayEventRatio": null,
                  "bandwidthLowEventRatio": null,
                  "networkTransportProtocol": "udp",
                  "traceRouteHops": []
                },
                "callerDevice": {
                  "captureDeviceName": "Realtek(R) Audio",
                  "captureDeviceDriver": "Realtek Semiconductor Corp.: 6.0.9780.1",
                  "renderDeviceName": "Realtek(R) Audio",
                  "renderDeviceDriver": "Realtek Semiconductor Corp.: 6.0.9780.1",
                  "sentSignalLevel": -23,
                  "receivedSignalLevel": null,
                  "sentNoiseLevel": -70,
                  "receivedNoiseLevel": null,
                  "initialSignalLevelRootMeanSquare": null,
                  "cpuInsufficentEventRatio": 0,
                  "renderNotFunctioningEventRatio": 0,
                  "captureNotFunctioningEventRatio": 0,
                  "deviceGlitchEventRatio": 0,
                  "lowSpeechToNoiseEventRatio": 0,
                  "lowSpeechLevelEventRatio": 0,
                  "deviceClippingEventRatio": 0,
                  "howlingEventCount": 0,
                  "renderZeroVolumeEventRatio": 0,
                  "renderMuteEventRatio": 0,
                  "micGlitchRate": null,
                  "speakerGlitchRate": null
                },
                "calleeDevice": {
                  "captureDeviceName": null,
                  "captureDeviceDriver": null,
                  "renderDeviceName": null,
                  "renderDeviceDriver": null,
                  "sentSignalLevel": null,
                  "receivedSignalLevel": null,
                  "sentNoiseLevel": null,
                  "receivedNoiseLevel": null,
                  "initialSignalLevelRootMeanSquare": null,
                  "cpuInsufficentEventRatio": null,
                  "renderNotFunctioningEventRatio": null,
                  "captureNotFunctioningEventRatio": null,
                  "deviceGlitchEventRatio": null,
                  "lowSpeechToNoiseEventRatio": null,
                  "lowSpeechLevelEventRatio": null,
                  "deviceClippingEventRatio": null,
                  "howlingEventCount": 0,
                  "renderZeroVolumeEventRatio": null,
                  "renderMuteEventRatio": null,
                  "micGlitchRate": null,
                  "speakerGlitchRate": null
                },
                "streams": [
                  {
                    "streamId": "648",
                    "startDateTime": null,
                    "endDateTime": null,
                    "streamDirection": "callerToCallee",
                    "averageAudioDegradation": 0,
                    "averageJitter": "PT0.002S",
                    "maxJitter": "PT0.01S",
                    "averagePacketLossRate": 0,
                    "maxPacketLossRate": 0,
                    "averageRatioOfConcealedSamples": 0.008604,
                    "maxRatioOfConcealedSamples": null,
                    "averageRoundTripTime": "PT0.047S",
                    "maxRoundTripTime": "PT0.049S",
                    "packetUtilization": 529,
                    "averageBandwidthEstimate": 2477474,
                    "wasMediaBypassed": null,
                    "postForwardErrorCorrectionPacketLossRate": null,
                    "averageVideoFrameLossPercentage": null,
                    "averageReceivedFrameRate": null,
                    "lowFrameRateRatio": null,
                    "averageVideoPacketLossRate": null,
                    "averageVideoFrameRate": null,
                    "lowVideoProcessingCapabilityRatio": null,
                    "averageAudioNetworkJitter": "PT0.008S",
                    "maxAudioNetworkJitter": "PT0.029S",
                    "audioCodec": "satin",
                    "videoCodec": "unknown",
                    "rmsFreezeDuration": null,
                    "averageFreezeDuration": null,
                    "isAudioForwardErrorCorrectionUsed": false
                  },
                  {
                    "streamId": "1000",
                    "startDateTime": null,
                    "endDateTime": null,
                    "streamDirection": "calleeToCaller",
                    "averageAudioDegradation": null,
                    "averageJitter": null,
                    "maxJitter": null,
                    "averagePacketLossRate": null,
                    "maxPacketLossRate": null,
                    "averageRatioOfConcealedSamples": null,
                    "maxRatioOfConcealedSamples": null,
                    "averageRoundTripTime": null,
                    "maxRoundTripTime": null,
                    "packetUtilization": 0,
                    "averageBandwidthEstimate": null,
                    "wasMediaBypassed": null,
                    "postForwardErrorCorrectionPacketLossRate": null,
                    "averageVideoFrameLossPercentage": null,
                    "averageReceivedFrameRate": null,
                    "lowFrameRateRatio": null,
                    "averageVideoPacketLossRate": null,
                    "averageVideoFrameRate": null,
                    "lowVideoProcessingCapabilityRatio": null,
                    "averageAudioNetworkJitter": null,
                    "maxAudioNetworkJitter": null,
                    "audioCodec": "unknown",
                    "videoCodec": "unknown",
                    "rmsFreezeDuration": null,
                    "averageFreezeDuration": null,
                    "isAudioForwardErrorCorrectionUsed": null
                  }
                ]
              }
            ]
          }
        ]
      }
    ],
    "participants_v2": [
      {
        "id": "42559007-03c6-42c8-971f-cb79fd381a5a",
        "identity": {
          "endpointType": null,
          "acsUser": null,
          "spoolUser": null,
          "phone": null,
          "guest": null,
          "encrypted": null,
          "onPremises": null,
          "acsApplicationInstance": null,
          "spoolApplicationInstance": null,
          "applicationInstance": null,
          "application": null,
          "device": null,
          "azureCommunicationServicesUser": null,
          "assertedIdentity": null,
          "user": {
            "id": "42559007-03c6-42c8-971f-cb79fd381a5a",
            "displayName": "John Doe",
            "tenantId": "9d66187e-13f0-4666-9bac-be67ddd4b676",
            "userPrincipalName": "john.doe@contoso.com"
          }
        },
        "administrativeUnitInfos": [
          {
            "id": "6d3eb178-988b-430b-974f-1254113e4522"
          }
        ]
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;
  let assertAccessTokenTypeStub: sinon.SinonStub;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');

    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');

    assertAccessTokenTypeStub = sinon.stub(accessToken, 'assertAccessTokenType').resolves();
  });

  afterEach(() => {
    sinonUtil.restore([
      accessToken.assertAccessTokenType,
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CALLRECORD_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when id is not a valid guid', async () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'invalid'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation when id is a valid guid', async () => {
    const actual = commandOptionsSchema.safeParse({
      id: validId
    });
    assert.strictEqual(actual.success, true);
  });

  it('retrieves the call record', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/communications/callRecords/${validId}?$expand=sessions($expand=segments)`) {
        return responseWithSessions;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/communications/callRecords/${validId}?$select=id&$expand=participants_v2`) {
        return responseWithParticipants;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: validId, verbose: true } });
    assert(loggerLogSpy.calledOnceWithExactly(response));
    assert(assertAccessTokenTypeStub.calledOnceWith('application'));
  });

  it('handles random API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'get').rejects(new Error(errorMessage));

    await assert.rejects(command.action(logger, { options: { id: 'validId' } }), new CommandError(errorMessage));
  });
});
