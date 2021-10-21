export const CanvasContentWebPart = {
  position: {
    zoneIndex: 1,
    sectionIndex: 1,
    controlIndex: 1,
    layoutIndex: 1
  },
  controlType: 3,
  id: 'af92a21f-a0ec-4668-ba2c-951a2b5d6f94',
  webPartId: 'cb3bfe97-a47f-47ca-bffb-bb9a5ff83d75',
  emphasis: {},
  zoneGroupMetadata: {
    type: 0
  },
  reservedHeight: 539,
  reservedWidth: 1180,
  addedFromPersistedData: true,
  webPartData: {
    id: 'cb3bfe97-a47f-47ca-bffb-bb9a5ff83d75',
    instanceId: 'af92a21f-a0ec-4668-ba2c-951a2b5d6f94',
    title: 'Conversations',
    description:
      'Show conversations from a Yammer group, user, topic, or home.',
    audiences: [],
    serverProcessedContent: {
      htmlStrings: {},
      searchablePlainTexts: {},
      imageSources: {},
      links: {}
    },
    dataVersion: '1.0',
    properties: {
      type: 'Home',
      showPublisher: true
    }
  }
};

export const mockControlGetData = {
  "CanvasContent1": JSON.stringify([{ ...CanvasContentWebPart }])
};

export const mockControlGetDataOutput = {
  id: 'af92a21f-a0ec-4668-ba2c-951a2b5d6f94',
  type: 'Client-side web part',
  title: 'Conversations',
  controlType: 3,
  order: 1,
  controlData: {
    position: {
      zoneIndex: 1,
      sectionIndex: 1,
      controlIndex: 1,
      layoutIndex: 1
    },
    controlType: 3,
    id: 'af92a21f-a0ec-4668-ba2c-951a2b5d6f94',
    webPartId: 'cb3bfe97-a47f-47ca-bffb-bb9a5ff83d75',
    emphasis: {},
    zoneGroupMetadata: {
      type: 0
    },
    reservedHeight: 539,
    reservedWidth: 1180,
    addedFromPersistedData: true,
    webPartData: {
      id: 'cb3bfe97-a47f-47ca-bffb-bb9a5ff83d75',
      instanceId: 'af92a21f-a0ec-4668-ba2c-951a2b5d6f94',
      title: 'Conversations',
      description:
        'Show conversations from a Yammer group, user, topic, or home.',
      audiences: [],
      serverProcessedContent: {
        htmlStrings: {},
        searchablePlainTexts: {},
        imageSources: {},
        links: {}
      },
      dataVersion: '1.0',
      properties: {
        type: 'Home',
        showPublisher: true
      }
    }
  }
};


export const CanvasContentText = {
  controlType: 4,
  id: '1212fc8d-dd6b-408a-8d5d-9f1cc787efbb',
  position: {
    controlIndex: 2,
    sectionIndex: 1,
    sectionFactor: 12,
    zoneIndex: 1,
    layoutIndex: 1
  },
  addedFromPersistedData: true,
  emphasis: {},
  zoneGroupMetadata: {
    type: 0
  },
  innerHTML: '<p>This is some text.</p>'
};

export const mockControlGetDataWithText = {
  CanvasContent1: JSON.stringify([
    {
      ...CanvasContentText
    }
  ])
};

export const mockControlGetDataWithTextOutput = {
  id: '1212fc8d-dd6b-408a-8d5d-9f1cc787efbb',
  type: 'Client-side text',
  controlType: 4,
  order: 1,
  controlData: {
    controlType: 4,
    id: '1212fc8d-dd6b-408a-8d5d-9f1cc787efbb',
    position: {
      controlIndex: 2,
      sectionIndex: 1,
      sectionFactor: 12,
      zoneIndex: 1,
      layoutIndex: 1
    },
    addedFromPersistedData: true,
    emphasis: {},
    zoneGroupMetadata: {
      type: 0
    },
    innerHTML: '<p>This is some text.</p>'
  }
};

export const mockControlGetDataEmptyColumn = {
  "CanvasContent1": JSON.stringify([{
    "id":"88f7b5b2-83a8-45d1-bc61-c11425f233e3",
    "position": { 
      "zoneIndex":2,
      "sectionIndex":1,
      "sectionFactor":12,
      "layoutIndex":1
    },
    "emphasis":{}
  }])
};

export const mockControlGetDataEmptyColumnOutput = {"id":"88f7b5b2-83a8-45d1-bc61-c11425f233e3","type":"Empty column","order":1,"controlData":{"id":"88f7b5b2-83a8-45d1-bc61-c11425f233e3","position":{"zoneIndex":2,"sectionIndex":1,"sectionFactor":12,"layoutIndex":1},"emphasis":{}}};

export const mockControlGetDataWithUnknownType = {
  CanvasContent1: JSON.stringify([
    {
      ...CanvasContentWebPart,
      controlType: 5
    },
    {
      id: 'EMPTY_0',
      position: {
        zoneIndex: 2,
        sectionIndex: 1,
        sectionFactor: 12,
        layoutIndex: 1
      },
      emphasis: {}
    }
  ])
};

export const mockControlGetDataWithUnknownTypeOutput = {
  id: 'af92a21f-a0ec-4668-ba2c-951a2b5d6f94',
  type: '5',
  title: 'Conversations',
  controlType: 5,
  order: 1,
  controlData: {
    position: {
      zoneIndex: 1,
      sectionIndex: 1,
      controlIndex: 1,
      layoutIndex: 1
    },
    controlType: 5,
    id: 'af92a21f-a0ec-4668-ba2c-951a2b5d6f94',
    webPartId: 'cb3bfe97-a47f-47ca-bffb-bb9a5ff83d75',
    emphasis: {},
    zoneGroupMetadata: {
      type: 0
    },
    reservedHeight: 539,
    reservedWidth: 1180,
    addedFromPersistedData: true,
    webPartData: {
      id: 'cb3bfe97-a47f-47ca-bffb-bb9a5ff83d75',
      instanceId: 'af92a21f-a0ec-4668-ba2c-951a2b5d6f94',
      title: 'Conversations',
      description:
        'Show conversations from a Yammer group, user, topic, or home.',
      audiences: [],
      serverProcessedContent: {
        htmlStrings: {},
        searchablePlainTexts: {},
        imageSources: {},
        links: {}
      },
      dataVersion: '1.0',
      properties: {
        type: 'Home',
        showPublisher: true
      }
    }
  }
};

export const mockControlGetDataWithoutAnId = {
  "CanvasContent1": JSON.stringify([{
    "id": null,
    "position": { 
      "zoneIndex":2,
      "sectionIndex":1,
      "sectionFactor":12,
      "layoutIndex":1
    },
    "emphasis":{}
  }])
};