'use strict';

module.exports = {
  async up(queryInterface, Sequelize) {
    // 1. Seed Kategori (Makanan, Minuman, Snack)
    await queryInterface.bulkInsert('category', [
      { category_id: 1, category_name: 'Makanan' },
      { category_id: 2, category_name: 'Minuman' },
      { category_id: 3, category_name: 'Snack' }
    ], {});

    // 2. Seed Produk sesuai dengan Katalog Edja Corner
    await queryInterface.bulkInsert('product', [
      {
        product_id: 1,
        product_name: 'Es Teh Manis',
        price: 5000,
        stock: 50,
        category_id: 2, // Minuman
        image: 'es-teh-manis.jpg'
      },
      {
        product_id: 2,
        product_name: 'Es Kopi Susu Gula Aren',
        price: 12000,
        stock: 30,
        category_id: 2, // Minuman
        image: 'es-kopi.jpg'
      },
      {
        product_id: 3,
        product_name: 'Nasi Ayam Original',
        price: 13000,
        stock: 25,
        category_id: 1, // Makanan
        image: 'nasi-ayam-ori.jpg'
      },
      {
        product_id: 4,
        product_name: 'Nasi Ayam Blackpaper',
        price: 15000,
        stock: 20,
        category_id: 1, // Makanan
        image: 'nasi-ayam-blackpaper.jpg'
      }
    ], {});
  },

  async down(queryInterface, Sequelize) {
    await queryInterface.bulkDelete('product', null, {});
    await queryInterface.bulkDelete('category', null, {});
  }
};