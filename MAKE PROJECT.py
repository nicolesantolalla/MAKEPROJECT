for d in json_price_data:
            price.append({'name': d['name'], 'price': float(d['amazon_price']), 'url': d['amazon_url']})
            price.append({'name': d['name'], 'price': float(d['walmart_price']), 'url': d['walmart_url']})
            price.append({'name': d['name'], 'price': float(d['ebay_price']), 'url': d['ebay_url']})
            minPricedItem = min(price, key=lambda x: x['price'])
            print(minPricedItem)
            print('=================')
            price = []
