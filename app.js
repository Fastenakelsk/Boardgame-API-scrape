const axios = require('axios');
const exceljs = require('exceljs');

const workbook = new exceljs.Workbook();
workbook.creator = 'Kilian Fastenakels';
const worksheet = workbook.addWorksheet('My Sheet');

worksheet.state = 'visible';

worksheet.properties.rowCount = 100;
worksheet.properties.columnCount = 29;

worksheet.columns = setColumns();

for(let i = 1; i < 70; i++){
  setTimeout(() => call('https://www.boardgameatlas.com/api/search?random=true&client_id=JLBr5npPhV', i, worksheet), 1000 * i);
  //setTimeout(() => call('https://www.boardgameatlas.com/api/search?name=Ticket to Ride: Nordic Countries&client_id=JLBr5npPhV', i, worksheet), 1000 * i);
}

async function call(url, i, sheet) {
  await axios.get(url)
  .then(response => {
    const game = response.data.games[0];
    console.log(game);

    const row = worksheet.getRow(i);
    
    row.values = {
      id: game.id ? game.id : "EMPTY",
      name: game.name ? game.name : "EMPTY",
      yearPublished: game.year_published ? game.year_published : "EMPTY",
      minPlayers: game.min_players ? game.min_players : "EMPTY",
      maxPlayers: game.max_players ? game.max_players : "EMPTY",
      minPlaytime: game.min_playtime ? game.min_playtime : "EMPTY",
      maxPlaytime: game.max_playtime ? game.max_playtime : "EMPTY",
      minAge: game.min_age ? game.min_age : "EMPTY",
      description: game.description ? game.description : "EMPTY",
      descriptionPreview: game.description_preview ? game.description_preview : "EMPTY",
      image: game.image_url ? game.image_url : "EMPTY",
      price: game.price ? game.price : "EMPTY",
      msrp: game.msrp ? game.msrp : "EMPTY",
      primaryPublisher: game.primary_publisher ? game.primary_publisher : "EMPTY",
      userRatings: game.num_user_ratings ? game.num_user_ratings : "EMPTY",
      averageRating: game.average_user_rating ? game.average_user_rating : "EMPTY",
      officialUrl: game.official_url ? game.official_url : "EMPTY",

      rulesUrl: game.rules_url ? game.rules_url : "EMPTY",
      weight: game.weight_amount ? game.weight_amount : "EMPTY",
      weightUnits: game.weight_units ? game.weight_units : "EMPTY",
      sizeHeight: game.size_height ? game.size_height : "EMPTY",
      sizeDepth: game.size_depth ? game.size_depth : "EMPTY",
      sizeUnits: game.size_units ? game.size_units : "EMPTY",

      historicalLowPrice: game.historical_low_price ? game.historical_low_price : "EMPTY",
      rank: game.rank ? game.rank : "EMPTY",
      trendingRank: game.trending_rank ? game.trending_rank : "EMPTY"
    };
    //publisher, designers, developers, artists, names
    //Remove all rows with empty values ?

    row.commit();

    workbook.csv.writeFile('test.csv');
    workbook.csv.writeFile('test.txt');
  })
  .catch(error => {
    console.log(error);
  });
}

function setColumns() {
  return [
    { header: 'id', key: 'id', width: 3 },
    { header: 'name', key: 'name', width: 100 },
    { header: 'yearPublished', key: 'yearPublished', width: 50 },
    { header: 'minPlayers', key: 'minPlayers', width: 50 },
    { header: 'maxPlayers', key: 'maxPlayers', width: 50 },
    { header: 'minPlaytime', key: 'minPlaytime', width: 50 },
    { header: 'maxPlaytime', key: 'maxPlaytime', width: 50 },
    { header: 'minAge', key: 'minAge', width: 50 },
    { header: 'description', key: 'description', width: 300 },
    { header: 'descriptionPreview', key: 'descriptionPreview', width: 300 },
    { header: 'image', key: 'image', width: 50 },
    { header: 'price', key: 'price', width: 50 },
    { header: 'msrp', key: 'msrp', width: 50 },
    { header: 'primaryPublisher', key: 'primaryPublisher', width: 50 },
    { header: 'publishers', key: 'publishers', width: 50 },
    { header: 'designers', key: 'designers', width: 50 },
    { header: 'developers', key: 'developers', width: 50 },
    { header: 'artists', key: 'artists', width: 50 },
    { header: 'names', key: 'names', width: 50 },
    { header: 'userRatings', key: 'userRatings', width: 50 },
    { header: 'averageRating', key: 'averageRating', width: 50 },
    { header: 'officialUrl', key: 'officialUrl', width: 50 },
    { header: 'rulesUrl', key: 'rulesUrl', width: 50 },
    { header: 'weight', key: 'weight', width: 50 },
    { header: 'weightUnits', key: 'weightUnits', width: 50 },
    { header: 'sizeHeight', key: 'sizeHeight', width: 50 },
    { header: 'sizeDepth', key: 'sizeDepth', width: 50 },
    { header: 'sizeUnits', key: 'sizeUnits', width: 50 },
    { header: 'historicalLowPrice', key: 'historicalLowPrice', width: 50 },
    { header: 'rank', key: 'rank', width: 50 },
    { header: 'trendingRank', key: 'trendingRank', width: 50 },
  ];
}