import { TeamsActivityHandler, CardFactory } from 'botbuilder';
import { getProductByName } from '../server/northwindDataService.js';

export class StockManagerBot extends TeamsActivityHandler {
    async handleTeamsMessagingExtensionQuery(context, query) {
        const { name, value } = query.parameters[0];
        if (name !== 'productName') {
            return;
        }

        const products = await getProductByName(value);
        const attachments = [];

        for (const pdt of products) {
            const heroCard = CardFactory.heroCard(pdt.productName);
            const preview = CardFactory.heroCard(pdt.productName);
            preview.content.tap = { type: 'invoke', value: { name: pdt.productName } };
            const attachment = { ...heroCard, preview };
            attachments.push(attachment);
        }

        var result = {
            composeExtension: {
                type: "result",
                attachmentLayout: "list",
                attachments: attachments
            }
        };

        return result;

    }
    async handleTeamsMessagingExtensionSelectItem(context, pdt) {
        return {
            composeExtension: {
                type: 'result',
                attachmentLayout: 'list',
                attachments: [CardFactory.heroCard(pdt.name)]
            }
        };

    }
}


