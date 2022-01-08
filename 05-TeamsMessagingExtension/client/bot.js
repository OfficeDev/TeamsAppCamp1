import { TeamsActivityHandler, CardFactory } from 'botbuilder';
import { getProductByName } from '../server/northwindDataService.js';
import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const pdtCardPayload =require('../client/cards/productCard.json');
import * as ACData from "adaptivecards-templating";
import * as AdaptiveCards from "adaptivecards";
export class StockManagerBot extends TeamsActivityHandler {
    async handleTeamsMessagingExtensionQuery(context, query){
        const { name, value } = query.parameters[0];
        if (name !== 'productName') {
            return;
        }

        const products = await getProductByName(value);
        const attachments = [];

        for (const pdt of products) {
            const heroCard = CardFactory.heroCard(pdt.productName);
            const preview = CardFactory.heroCard(pdt.productName);
            preview.content.tap = { type: 'invoke', value: { name: pdt.productName,
            id:pdt.productId } };
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
        const preview = CardFactory.heroCard(pdt.name);
        var template = new ACData.Template(pdtCardPayload);
      
        var card = template.expand({
            $root: {
                pdt
            }
        });
        var adaptiveCard = new AdaptiveCards.AdaptiveCard();
        adaptiveCard.parse(card);
        const adaptive = CardFactory.adaptiveCard(card);
        const attachment = { ...adaptive, preview };
        return {
            composeExtension: {
                type: 'result',
                attachmentLayout: 'list',
                attachments: [attachment]
            },
        };
    }   
    
}
    
    



